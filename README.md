# Tablero local de seguimiento de préstamos

## Qué hace
- Corre en tu PC con `http://127.0.0.1:5000`.
- Lee el Excel desde `data/Prestamos 04-2026 - 03-2027.xlsm`.
- Usa la hoja `Cuotas` para armar indicadores.
- Tiene login y creación de usuarios.
- Es solo visualización: no modifica el Excel.

## Usuario inicial
- Usuario: `admin`
- Contraseña: `admin123`

## Instalación en Windows
1. Instalar Python 3.11 o superior.
2. Abrir CMD o PowerShell en esta carpeta.
3. Ejecutar:

```bat
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
python app.py
```

4. Abrir en el navegador:

```text
http://127.0.0.1:5000
```

## Actualización desde Excel
Cada vez que guardes cambios en el Excel, entrá al tablero o tocá “Actualizar desde Excel”. La app vuelve a leer la hoja `Cuotas`.

## Cambiar archivo Excel
Editá el archivo `.env` y cambiá la línea:

```text
EXCEL_PATH=data/Prestamos 04-2026 - 03-2027.xlsm
```

Podés poner una ruta absoluta, por ejemplo:

```text
EXCEL_PATH=C:\Users\Usuario\Desktop\Prestamos.xlsm
```

## Más adelante: acceso por red interna
Para que otras PCs vean el tablero, se cambia en `app.py`:

```python
app.run(host="0.0.0.0", port=5000, debug=False)
```

y se abre el puerto 5000 en Firewall de Windows. Luego acceden con:

```text
http://IP-DE-TU-PC:5000
```


## Cambios v5
- La sección Cuotas incorpora filtro por período.
- Cruza Cuotas con Seguimiento por ID del préstamo.
- Agrega columnas: cuotas pendientes, TNA, TEA, fecha de toma y tipo de amortización.
- Mantiene Dashboard sin cambios funcionales.
