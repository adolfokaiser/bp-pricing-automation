"""
Script para copiar un archivo de plantilla de Excel
y generar múltiples copias con nombres personalizados.
Las rutas de origen y destino se leen desde un archivo .env.
"""

import os
import shutil
from dotenv import load_dotenv

# ─── Cargar configuración desde .env ─────────────────────────────
load_dotenv()

TEMPLATE_PATH = os.getenv("TEMPLATE_PATH")    # Ruta al archivo Excel plantilla
DEST_FOLDER   = os.getenv("DEST_FOLDER")      # Carpeta donde se guardarán las copias

# ─── Lista de nombres personalizados para las copias ─────────────
NAMES = [
    "BP Buenavista Appro",
    "BP Casa Blanca (La Diez)",
    "BP Diaz Ordaz (agil)",
    "BP Galeana (agil)",
    "BP Insurgentes (La Diez)",
    "BP Juan Ruiz de Alarcon (Appro)",
    "BP La Diez (La Diez)",
    "BP Otay (Appro)",
    "BP Peñasco (Appro)",
    "BP Puente Machado (agil)",
    "BP Rancho Viejo",
    "BP UABC (agil)",
    "BP Contadero",
    "BP Eje 3",
    "BP Felix Cuevas",
    "BP Miramontes",
    "BP Leon I",
    "BP Chapultepec",
    "BP Cuautitlan Centro",
    "BP Fresnos",
    "BP Lago de Guadalupe I (Micha)",
    "BP Lago de Guadalupe II (Micha)",
    "BP Melchor Ocampo",
    "BP Orquidea",
    "BP Pantitlan",
    "BP Perinorte",
    "BP Puente de Vigas",
    "BP Santa Cecilia",
    "BP Satelite",
    "BP Tecnologico",
    "BP Camino al Batan",
    "BP Periferico Ecologico",
    "BP Tequisquiapan",
    "BP Cancun- Tulum",
    "BP Apizaco (Andariego)",
    "BP Coatza Mina",
    "BP Juan Escutia",
    "BP Llano de en medio",
    "BP Nueva Obrera",
    "BP Tihuatlan"
]

def main():
    # Verificar que las variables de entorno estén definidas
    if not TEMPLATE_PATH or not DEST_FOLDER:
        print("❌ Error: define TEMPLATE_PATH y DEST_FOLDER en el archivo .env")
        return

    # Crear la carpeta de destino si no existe
    os.makedirs(DEST_FOLDER, exist_ok=True)

    # Copiar el archivo plantilla con cada nombre
    for name in NAMES:
        dest_filename = f"{name} P.xlsx"
        dest_path = os.path.join(DEST_FOLDER, dest_filename)
        try:
            shutil.copy(TEMPLATE_PATH, dest_path)
        except Exception as e:
            print(f"❌ Error al copiar para '{name}': {e}")

    print(f"✅ Se copiaron {len(NAMES)} archivos a:\n   {DEST_FOLDER}")

if __name__ == "__main__":
    main()
