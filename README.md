# SGSI Ciberseguridad – Sistema de Gestión de Seguridad de la Información

Este repositorio contiene plantillas, documentos y automatizaciones para un SGSI (Sistema de Gestión de Seguridad de la Información) con soporte para Excel con macros (XLSM), generación de documentos Word y utilidades avanzadas de tuning de riesgos.

## Contenido principal

### Archivo XLSM principal
- **`SGSI_COMPLETO_David_Nanculeo_v5.0.xlsm`**: Versión operativa del SGSI con todas las hojas configuradas (Matriz_Riesgos, Panel_Control, Dashboard, Config_Tuning, etc.)

### Macros VBA
- **`SGSI_MACROS_VBA_COMPLETAS_v5.0_CONSOLIDADO.txt`**: Colección completa de macros VBA para el SGSI
- **`SGSI_MACRO_TUNING_AUDITOR.txt`**: Macro especializada para aplicar tuning del auditor (ajuste experto de riesgos)

### Scripts de Tuning (Python)
- **`tuning_auditor_sgsi.py`**: Script principal de tuning para ajustar riesgos con criterio experto
- **`tuning_auditor.py`**: Versión alternativa del script de tuning

### Documentación
- **`Tuning.pdf`**: Guía PDF sobre el sistema de tuning del auditor
- **`documentacion/`**: Carpeta con material de apoyo, plantillas y presentaciones

### Generadores de documentos Word (SGSI Premium)
Scripts que crean documentos `.docx` con estructura profesional:
- `00_Procedimiento...py`, `01_Plan_de_proyecto...py`, `02_Acta_de_constitucion...py`
- Scripts genéricos en `generar_*`

### Referencias legales y compliance
- `Literatura Compliance (1).pdf`
- Legislación chilena: Ley 21459, Ley 21663, Ley 21719

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

## Sistema de Tuning del Auditor

### ¿Qué es el Tuning?
El tuning o evaluación del experto es un mecanismo que permite al auditor ajustar el valor de riesgo calculado automáticamente (probabilidad × impacto) cuando identifica factores contextuales no capturados por el modelo base:
- Controles de seguridad no formalizados pero efectivos
- Exposición real diferente a la estimada
- Obligaciones regulatorias o contractuales específicas
- Criticidad estratégica del proceso o activo

### Implementación dual (VBA + Python)

**Opción A: Macro VBA** (para uso interactivo en Excel)
1. Abre `SGSI_COMPLETO_David_Nanculeo_v5.0.xlsm`
2. Ejecuta la macro `AplicarTuningAuditor` desde el Panel_Control o Editor VBA (Alt+F11)
3. La macro crea/actualiza:
   - Hoja `Config_Tuning` con escala 1–5 y factores (0.70–1.30)
   - Columnas en `Matriz_Riesgos`: `TUNING_AUDITOR`, `FACTOR_TUNING`, `RIESGO_TUNING`, `NIVEL_TUNING`
   - Formato condicional de color según criticidad (BAJO/MEDIO/ALTO/CRÍTICO)

**Opción B: Script Python** (para automatización y procesamiento masivo)
```powershell
python tuning_auditor_sgsi.py "SGSI_COMPLETO_David_Nanculeo_v5.0.xlsm"
```

**Salida:**
- Genera archivo con sufijo `_TUNING` preservando todas las macros
- Ejemplo: `SGSI_COMPLETO_David_Nanculeo_v5.0_TUNING.xlsm`

### Escala de Tuning

| Nivel | Descripción | Factor | Efecto |
|-------|-------------|--------|--------|
| 1 | Muy por debajo de la estimación | 0.70 | Reduce riesgo 30% |
| 2 | Por debajo de la estimación | 0.85 | Reduce riesgo 15% |
| 3 | Confirma la estimación (neutro) | 1.00 | Sin cambios |
| 4 | Por encima de la estimación | 1.15 | Incrementa 15% |
| 5 | Criticidad máxima confirmada | 1.30 | Incrementa 30% o fuerza a 25 |

**Regla crítica:** Cuando el experto asigna nivel 5, el sistema puede forzar el valor a 25 (techo de la escala) para garantizar visibilidad en tableros de gestión.

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

## Archivos de Presentación y Documentación

### Presentación del Sistema de Tuning
- **`documentacion/presentacion.md`**: Guía explicativa completa (10 diapositivas) sobre el sistema de tuning
  - Definición y objetivo del tuning
  - Arquitectura VBA + Python
  - Configuración de la escala
  - Columnas requeridas y cálculos
  - Normalización y factor multiplicativo
  - Regla crítica para nivel 5
  - Clasificación y formato visual
  - Instrucciones de ejecución paso a paso
  - Verificación y KPIs
  - Solución de problemas y estrategia de pruebas

### Documentación Técnica
Consulta `Tuning.pdf` para la guía completa sobre:
- Fundamentos teóricos del tuning del auditor
- Casos de uso y ejemplos prácticos
- Integración con Panel_Control y Dashboard
- Buenas prácticas de asignación de niveles

## Buenas prácticas
- ✓ Crear respaldo del archivo antes de aplicar tuning masivo
- ✓ Trabajar sobre la versión operativa (`v5.0.xlsm`)
- ✓ Mantener nombres estándar de hojas clave (`Matriz_Riesgos`, `Config_Tuning`)
- ✓ Evitar ediciones manuales en encabezados combinados
- ✓ Documentar criterios de asignación de niveles 4 y 5 para auditoría posterior
- ✓ Validar que `Panel_Control`/`Dashboard` reflejen correctamente los valores ajustados

## Estrategia de pruebas recomendada
1. **Fase 1:** Carga 3–5 riesgos de prueba con valores conocidos
2. **Fase 2:** Asigna niveles variados (1, 3, 4, 5) para observar el comportamiento
3. **Fase 3:** Verifica que nivel 5 con riesgo base > 0 se fuerza a 25
4. **Fase 4:** Valida colores y clasificaciones (BAJO/MEDIO/ALTO/CRÍTICO)
5. **Fase 5:** Comprueba que KPIs en Panel_Control reflejan los nuevos valores

## Estructura adicional
- `documentacion/`: Material de apoyo, plantillas y presentaciones
- `Plantilla_Documentos_SGSI_Fase1*.docx`: Base para nuevos documentos

## Soporte
Si detectas errores o necesitas ampliar funcionalidades (más hojas, integración con DRP/Mantenimiento, nueva matriz de riesgos, etc.), abre un issue o indica los requisitos para extender los scripts.
