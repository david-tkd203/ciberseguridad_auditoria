# SGSI Ciberseguridad – README

Este repositorio contiene plantillas, documentos y automatizaciones para un SGSI (Sistema de Gestión de Seguridad de la Información) con soporte para Excel con macros (XLSM), generación de documentos Word y utilidades de tuning de riesgos.

## Contenido principal
- Archivos XLSM del SGSI (múltiples versiones): `SGSI_COMPLETO_David_Nanculeo_v4.0*.xlsm`, `v4.1`, `v4.2`, `v4.5_FINAL`.
- Macros VBA consolidadas: `SGSI_MACROS_VBA_COMPLETAS_v5.0_CONSOLIDADO.txt` y módulo específico `SGSI_MACRO_TUNING_AUDITOR.txt`.
- Utilidad de Tuning (Python) para Matriz de Riesgos: `tuning_auditor.py`.
- Generadores de documentos Word (python-docx): `00_Procedimiento...py`, `01_Plan_de_proyecto...py`, `02_Acta_de_constitucion...py`, y scripts genéricos en `generar_*`.
- Referencias legales y compliance: `Literatura Compliance (1).pdf`, `Ley-21459_20-JUN-2022.pdf`, `Ley-21663_08-ABR-2024.pdf`, `Ley-21719_13-DIC-2024-2.pdf`.

## Documentos SGSI habilitados
Incluye la siguiente tabla de identificación de documentos estándar del SGSI:

| Identificador | Nombre Documento                                             | Versión | Dueño |
|---------------|---------------------------------------------------------------|---------|-------|
| PSI-001-NN    | Política de Seguridad de la Información                      | 1.x     | -     |
| POL-PDR-002   | Procedimiento para el Control de Documentos y Registros      | 1.x     | -     |
| PLAN-PRO-003  | Plan de Proyecto SGSI                                        | 1.x     | -     |
| ALC-SGSI-004  | Alcance del SGSI                                             | 1.x     | -     |
| MET-RIS-005   | Metodología de Evaluación y Tratamiento de Riesgos           | 1.x     | -     |
| PROC-SOA-006  | Declaración de Aplicabilidad SoA                             | 1.x     | -     |
| PLAN-MIT-007  | Plan Mitigación de Riesgos                                   | 1.x     | -     |
| PLAN-DIR-008  | Plan Director de Ciberseguridad                              | 1.x     | -     |
| PROC-NDA-009  | Cláusulas de Seguridad para Proveedores (NDA)                | 1.x     | -     |
| PROC-DCP-010  | Declaración de Confidencialidad y Privacidad (NDA) - Empleados| 1.x     | -     |
| INST-TCC-011  | Actas Tipo Comité de Crisis                                  | 1.x     | -     |
| PLAN-FOR-012  | Plan de Formación y Concienciación                           | 1.x     | -     |
| PROC-IAI-013  | Inventario de Activos de Información y CIA                   | 1.x     | -     |
| PROC-CMDB-014 | Procedimiento de Gestión de Configuración (CMDB)             | 1.x     | -     |
| PROC-HAR-015  | Procedimiento de Hardening                                   | 1.x     | -     |
| PROC-GIP-016  | Gestión de Incidentes: Procedimiento y Registro              | 1.x     | -     |
| POL-CDN-017   | Continuidad del Negocio                                      | 1.x     | -     |
| POL-CON-018   | Política de Continuidad                                      | 1.x     | -     |
| MET-BIA-019   | Metodología y Cuestionario BIA                               | 1.x     | -     |
| PLAN-PPV-020  | Plan de Pruebas y Verificación                               | 1.x     | -     |
| INST-AUD-021  | Auditoría Interna                                            | 1.x     | -     |
| INST-IMA-022  | Informe de Medición y Actas de Revisión                      | 1.x     | -     |

## Requisitos
- Windows con Microsoft Excel (para abrir `.xlsm`).
- Python 3.9+.
- Paquetes Python:
  - `openpyxl` (manejo de Excel y macros – lectura y escritura; mantiene `keep_vba=True`).
  - `python-docx` (generación de documentos Word).

Instalación rápida (PowerShell):
```powershell
pip install openpyxl python-docx
```

## Uso de la utilidad de Tuning del Auditor
`tuning_auditor.py` aplica una evaluación experta (TUNING) sobre la hoja `Matriz_Riesgos` del archivo SGSI, creando/actualizando:
- `Config_Tuning` con escala 1–5 y factores (0.70–1.30).
- Columnas en `Matriz_Riesgos`: `TUNING_AUDITOR`, `FACTOR_TUNING`, `RIESGO_TUNING`, `NIVEL_TUNING`.
- Formato condicional de color según criticidad.

Ejemplo de ejecución (usa por defecto `SGSI_COMPLETO_David_Nanculeo_v4.5_FINAL.xlsm`):
```powershell
python tuning_auditor.py
```
O indicando un archivo específico:
```powershell
python tuning_auditor.py SGSI_COMPLETO_David_Nanculeo_v4.1.xlsm
```
Salida:
- Guarda un archivo nuevo con sufijo `_TUNING` (p. ej. `SGSI_COMPLETO_David_Nanculeo_v4.5_FINAL_TUNING.xlsm`).

## Generación de documentos Word (SGSI Premium)
Scripts que crean documentos `.docx` con estructura profesional:
- `00_Procedimiento_para_el_control_de_documentos_y_registros_Premium_ES.py`
- `01_Plan_de_proyecto_Premium_ES.py`
- `02_Acta_de_constitucion_del_proyecto_Premium_ES.py`

Ejecutar, por ejemplo:
```powershell
python 00_Procedimiento_para_el_control_de_documentos_y_registros_Premium_ES.py
python 01_Plan_de_proyecto_Premium_ES.py
python 02_Acta_de_constitucion_del_proyecto_Premium_ES.py
```
Los documentos se generan en la carpeta raíz del proyecto.

## Referencias normativas y compliance
- ISO 27001 (Seguridad de la Información) y ISO 22301 (Continuidad de Negocio).
- Literatura Compliance (1).pdf.
- Legislación: Ley 21459, Ley 21663, Ley 21719.

Cada documento y herramienta procura alinear controles, procesos y evidencias con estas referencias.

## Buenas prácticas
- Mantener una copia de respaldo del `.xlsm` antes de ejecutar automatizaciones.
- Trabajar sobre la versión más reciente del archivo (`v4.5_FINAL`).
- Validar que la hoja `Matriz_Riesgos` y los encabezados estén presentes.
- Revisar el resultado `_TUNING` y ajustar `TUNING_AUDITOR` (1–5) por fila donde aplique.

## Estructura adicional
- `documentacion/` y `documentacion.zip`: material de apoyo y plantillas.
- `Plantilla_Documentos_SGSI_Fase1*.docx`: base para nuevos documentos.

## Soporte
Si detectas errores o necesitas ampliar funcionalidades (más hojas, integración con DRP/Mantenimiento, nueva matriz de riesgos, etc.), abre un issue o indica los requisitos para extender los scripts.
