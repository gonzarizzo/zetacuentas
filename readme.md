## Conda environment
1. Create it (once): `mamba env create -f environment.yml`
2. Use it: `mamba run -n mi_entorno python <script.py>`

## Load Itaú Crédito data
1. Copy the credit-card movement table from Itaú and paste it into Excel.
2. Save that Excel as `movimientos.csv`.
3. Run `mamba run -n mi_entorno python generar_excels.py` to build `movimientos_pesos.xlsx` and `movimientos_dolares.xlsx`.

## Load Itaú Débito data
1. Download the current-month movements as Excel (bottom of the list in Itaú web).
2. Place the files in this folder.
3. Run `mamba run -n mi_entorno python procesar_estados_itau.py` to produce `itau_debito_pesos.xlsx` and `itau_debito_dolares.xlsx`.

## Load BROU data
1. Download the current-month movements as XLS from the “Guardar -> Archivo XLS” button (top-right of the BROU page).
2. Place the file as `Detalle_Movimiento_Cuenta.xls` in this folder.
3. Run `mamba run -n mi_entorno python procesar_movimiento_brou.py` to create `brou_detalle_movimientos.xlsx`.

## Remove movements already recorded
1. Go to zetacuenta, login, go to "Comprobantes -> Excel" and overwrite the content to `comprobante.xlsx`.
2. Use `filtrar_movimientos_registrados.py` whenever `comprobante.xlsx` has been updated. It removes any rows already present in the comprobante from:

- `movimientos_pesos.xlsx` (`Crédito Itaú $`)
- `movimientos_dolares.xlsx` (`Crédito Itaú U$S`)
- `brou_detalle_movimientos.xlsx` (`Débito BROU $`)
- `itau_debito_pesos.xlsx` (`Débito Itaú $ Gonza`)

Run it with:

```
mamba run -n mi_entorno python filtrar_movimientos_registrados.py
```

The script overwrites the Excel files above (creating backups is recommended if you need to keep the originals).

3. Upload the Excel files of each bank account to Zetacuentas and complete the missing info.
