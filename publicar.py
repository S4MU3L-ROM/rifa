"""Lee rifa_celular.xlsx y publica datos.json (sin nombres ni teléfonos) en GitHub Pages."""
import json
import subprocess
from pathlib import Path

import openpyxl

DIR = Path(__file__).parent


def leer_excel(ruta):
    wb = openpyxl.load_workbook(ruta, data_only=True)
    cfg = {str(r[0].value).strip(): r[1].value for r in wb["Config"].iter_rows(min_row=4) if r[0].value}
    vendidos, apartados = [], []
    for num, _nombre, _tel, estado, *_ in wb["Boletas"].iter_rows(min_row=2, values_only=True):
        if num is None:
            continue
        n = str(num).zfill(2)
        estado = str(estado or "").strip().lower()
        if estado == "pagado":
            vendidos.append(n)
        elif estado == "apartado":
            apartados.append(n)
    fecha = cfg["Fecha del sorteo"]
    return {
        "premio": int(cfg["Valor del premio"]),
        "precio": int(cfg["Precio por boleta"]),
        "loteria": cfg["Lotería"],
        "sorteo": fecha.strftime("%Y-%m-%d"),
        "regla": cfg["Si el número ganador no se vendió"],
        "vendidos": vendidos,
        "apartados": apartados,
    }


def git(*args):
    return subprocess.run(["git", *args], cwd=DIR, capture_output=True, text=True)


if __name__ == "__main__":
    datos = leer_excel(DIR / "rifa_celular.xlsx")
    (DIR / "datos.json").write_text(json.dumps(datos, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Premio ${datos['premio']:,} | Sorteo {datos['sorteo']} | "
          f"Vendidos {len(datos['vendidos'])} | Apartados {len(datos['apartados'])}".replace(",", "."))

    git("add", "datos.json")
    if git("diff", "--cached", "--quiet").returncode == 0:
        print("No hay cambios: la página ya está al día.")
    else:
        git("commit", "-m", "Actualizar tablero desde Excel")
        r = git("push")
        print("¡Publicado! En 1-2 minutos se ve en https://s4mu3l-rom.github.io/rifa/" if r.returncode == 0
              else "ERROR al subir:\n" + r.stderr)
