#!/usr/bin/env python3
"""
Generador topográfico de la tortuga RMA (Vol. 20 · 2027).
Traza curvas de nivel (isolíneas) a partir del campo de distancia suavizado
del emblema de la tortuga sobre el fondo borravino patrimonial (#820D27).
"""
import argparse
import hashlib
import math
import random
from pathlib import Path


def smin(a: float, b: float, k: float = 28.0) -> float:
    h = max(k - abs(a - b), 0.0) / k
    return min(a, b) - h * h * k * 0.25


def smoothstep(edge0: float, edge1: float, x: float) -> float:
    t = max(0.0, min(1.0, (x - edge0) / (edge1 - edge0)))
    return t * t * (3.0 - 2.0 * t)


def generar_svg(vol: int, num: int, year: int, out_path: Path) -> str:
    seed_hex = hashlib.sha256(f"RMA-{vol}-{num}-{year}".encode()).hexdigest()[:12]
    rng = random.Random(int(seed_hex, 16))

    W, H = 900, 920
    cx = W * 0.50 + rng.uniform(-12, 12)
    cy = H * 0.54 + rng.uniform(-10, 10)

    # Parámetros de onda topográfica deterministas por fascículo
    p1 = rng.uniform(0, math.tau)
    p2 = rng.uniform(0, math.tau)
    p3 = rng.uniform(0, math.tau)
    f1 = rng.uniform(0.0085, 0.0115)
    f2 = rng.uniform(0.014, 0.019)

    # Silueta del glifo RMA (caparazón + cabeza + 4 extremidades + cola)
    r_shell = 116.0
    partes = [
        (0.0, -142.0, 38.0),    # Cabeza
        (-106.0, -82.0, 36.0),  # Extremidad superior izquierda
        (106.0, -82.0, 36.0),   # Extremidad superior derecha
        (-96.0, 92.0, 34.0),    # Extremidad inferior izquierda
        (96.0, 92.0, 34.0),     # Extremidad inferior derecha
        (0.0, 134.0, 21.0),     # Cola
    ]

    def campo(x: float, y: float) -> float:
        d = math.hypot(x - cx, y - cy) - r_shell
        for dx, dy, r in partes:
            dp = math.hypot(x - (cx + dx), y - (cy + dy)) - r
            d = smin(d, dp, 26.0)

        if d > 0.0:
            influencia = smoothstep(0.0, 165.0, d)
            onda = (
                34.0 * math.sin(x * f1 + y * 0.006 + p1)
                + 22.0 * math.cos(y * f2 - x * 0.007 + p2)
                + 14.0 * math.sin((x - y) * 0.012 + p3)
            )
            d += onda * influencia
        return d

    nx, ny = 180, 184
    dx_g = W / nx
    dy_g = H / ny
    grid = [[campo(ix * dx_g, iy * dy_g) for ix in range(nx + 1)] for iy in range(ny + 1)]

    def isolinea(nivel: float) -> str:
        segs = []
        for iy in range(ny):
            y0 = iy * dy_g
            y1 = y0 + dy_g
            row0 = grid[iy]
            row1 = grid[iy + 1]
            for ix in range(nx):
                v0 = row0[ix] - nivel
                v1 = row0[ix + 1] - nivel
                v2 = row1[ix + 1] - nivel
                v3 = row1[ix] - nivel

                idx = (1 if v0 > 0 else 0) | (2 if v1 > 0 else 0) | (4 if v2 > 0 else 0) | (8 if v3 > 0 else 0)
                if idx == 0 or idx == 15:
                    continue

                x0 = ix * dx_g
                x1 = x0 + dx_g

                def interp(xa, ya, va, xb, yb, vb):
                    denom = vb - va
                    t = 0.5 if abs(denom) < 1e-9 else -va / denom
                    return (xa + t * (xb - xa), ya + t * (yb - ya))

                top = interp(x0, y0, v0, x1, y0, v1) if (idx & 1) != ((idx >> 1) & 1) else None
                right = interp(x1, y0, v1, x1, y1, v2) if ((idx >> 1) & 1) != ((idx >> 2) & 1) else None
                bottom = interp(x0, y1, v3, x1, y1, v2) if ((idx >> 3) & 1) != ((idx >> 2) & 1) else None
                left = interp(x0, y0, v0, x0, y1, v3) if (idx & 1) != ((idx >> 3) & 1) else None

                if idx in (1, 14):
                    segs.append((top, left))
                elif idx in (2, 13):
                    segs.append((top, right))
                elif idx in (3, 12):
                    segs.append((left, right))
                elif idx in (4, 11):
                    segs.append((right, bottom))
                elif idx == 5:
                    segs.append((top, right))
                    segs.append((left, bottom))
                elif idx in (6, 9):
                    segs.append((top, bottom))
                elif idx in (7, 8):
                    segs.append((left, bottom))
                elif idx == 10:
                    segs.append((top, left))
                    segs.append((right, bottom))

        adj = {}
        def key(pt):
            return (round(pt[0], 2), round(pt[1], 2))

        for a, b in segs:
            ka, kb = key(a), key(b)
            adj.setdefault(ka, []).append(kb)
            adj.setdefault(kb, []).append(ka)

        visitados = set()
        paths = []
        for start in adj:
            if len(adj[start]) == 1 and start not in visitados:
                chain = [start]
                visitados.add(start)
                curr = start
                while True:
                    nxts = [n for n in adj[curr] if n not in visitados]
                    if not nxts:
                        break
                    curr = nxts[0]
                    visitados.add(curr)
                    chain.append(curr)
                if len(chain) > 1:
                    paths.append("M " + " L ".join(f"{x:.1f} {y:.1f}" for x, y in chain))

        for start in adj:
            if start not in visitados:
                chain = [start]
                visitados.add(start)
                curr = start
                while True:
                    nxts = [n for n in adj[curr] if n not in visitados]
                    if not nxts:
                        break
                    curr = nxts[0]
                    visitados.add(curr)
                    chain.append(curr)
                if len(chain) > 2:
                    paths.append("M " + " L ".join(f"{x:.1f} {y:.1f}" for x, y in chain) + " Z")

        return " ".join(paths)

    niveles_int = [-96.0, -76.0, -56.0, -36.0, -21.0, -9.5]
    niveles_ext = [i * 19.5 for i in range(1, 29)]

    capas = []
    for niv in niveles_int:
        d_str = isolinea(niv)
        if d_str:
            capas.append(f'<path d="{d_str}" fill="none" stroke="#F2DFC4" stroke-width="1.3" stroke-opacity="0.44"/>')

    # Contorno maestro de la tortuga
    d_zero = isolinea(0.0)
    if d_zero:
        capas.append(f'<path d="{d_zero}" fill="none" stroke="#F2DFC4" stroke-width="2.8" stroke-opacity="0.88"/>')

    # Curvas topográficas envolventes
    for idx, niv in enumerate(niveles_ext, start=1):
        d_str = isolinea(niv)
        if not d_str:
            continue
        if idx % 5 == 0:
            capas.append(f'<path d="{d_str}" fill="none" stroke="#4A0514" stroke-width="1.6" stroke-opacity="0.52"/>')
        else:
            op = 0.40 if idx <= 4 else 0.28
            capas.append(f'<path d="{d_str}" fill="none" stroke="#F2DFC4" stroke-width="1.15" stroke-opacity="{op:.2f}"/>')

    svg = f"""<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" width="{W}" height="{H}">
  <rect width="{W}" height="{H}" fill="#820D27"/>
  {chr(10) + "  ".join(capas)}
</svg>
"""
    out_path.write_text(svg, encoding="utf-8")
    return seed_hex


def main():
    parser = argparse.ArgumentParser(description="Genera trama topográfica de la tortuga RMA.")
    parser.add_argument("--vol", type=int, default=20)
    parser.add_argument("--num", type=int, default=1)
    parser.add_argument("--year", type=int, default=2027)
    parser.add_argument("--out", type=Path, default=Path("trama.svg"))
    args = parser.parse_args()

    seed = generar_svg(args.vol, args.num, args.year, args.out)
    print(f"[OK] Trama topográfica de tortuga generada en {args.out} | Vol. {args.vol} Núm. {args.num} ({args.year}) | Seed: {seed}")


if __name__ == "__main__":
    main()
