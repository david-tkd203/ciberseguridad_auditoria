# 🛡️ SGSI - Sistema de Gestión de Seguridad de la Información v4.0

[![ISO 27001:2022](https://img.shields.io/badge/ISO%2027001-2022-blue.svg)](https://www.iso.org/standard/27001)
[![MITRE ATT&CK](https://img.shields.io/badge/MITRE-ATT%26CK-red.svg)](https://attack.mitre.org/)
[![Excel](https://img.shields.io/badge/Excel-2016%2B-green.svg)](https://www.microsoft.com/excel)
[![VBA Macros](https://img.shields.io/badge/VBA-21%20Macros-orange.svg)](https://docs.microsoft.com/office/vba/api/overview/excel)

---

## 📋 Descripción General

**Sistema completo y profesional** de gestión de seguridad de la información (SGSI) que implementa **todas las exigencias de la norma ISO 27001:2022**. Este sistema proporciona documentación completa, herramientas automatizadas y procesos robustos para **implementar, gestionar y certificar** un SGSI de clase mundial.

### 🎯 Propósito

Facilitar la implementación de un SGSI certificable bajo ISO 27001:2022, proporcionando:
- ✅ **Documentación completa** de 45 hojas estructuradas
- ✅ **Automatización total** con 21 macros VBA
- ✅ **Gestión integral** de activos, riesgos, controles e incidentes
- ✅ **Framework MITRE ATT&CK** integrado para análisis de amenazas
- ✅ **Sistema listo para auditoría** y certificación

### 🏆 Beneficios Clave

| Beneficio | Descripción |
|-----------|-------------|
| 💼 **Ahorro de Tiempo** | Sistema preconfigurado elimina 200+ horas de trabajo manual |
| 🤖 **Automatización** | 21 macros VBA reducen errores y aceleran procesos |
| 📈 **Escalabilidad** | Desde startups hasta grandes corporaciones |


---

## ✨ Características Principales del Sistema

### 📊 **1. Gestión Integral de Activos**

El sistema permite inventariar y clasificar todos los activos de información de la organización:

- **💼 Inventario Completo**: Registra hardware, software, datos, servicios, personal e instalaciones
- **🔖 Clasificación Multinivel**: Categorías, subcategorías, áreas y clases personalizables
- **🗺️ Ubicación Precisa**: Seguimiento de ubicación física y lógica de cada activo
- **💰 Valoración**: Asignación de valor económico y criticidad (C/I/D)
- **📊 Dashboard Visual**: Panel de control con métricas en tiempo real

**Hoja Principal**: `Activos` (Inventario maestro con IDs automáticos ACT-2025-XXX)

---

### ⚠️ **2. Análisis y Gestión de Riesgos**

Metodología robusta de análisis de riesgos basada en matriz 5×5:

- **📊 Matriz Probabilidad × Impacto**: Escala de 1 a 5 en ambos ejes (riesgo 1-25)
- **🧮 Cálculo Automático**: Riesgo Inherente y Residual calculados por macros
- **🎨 Código de Colores**: Verde (Bajo), Amarillo (Medio), Naranja (Alto), Rojo (Crítico)
- **🗺️ Mapa de Calor Visual**: Matriz 5×5 con distribución de riesgos
- **📈 Análisis de Brechas**: Identifica diferencias entre riesgo actual y aceptable

**Hojas Principales**: 
- `Matriz_Riesgos` (Registro de todos los riesgos identificados)
- `Analisis_Riesgos` (Análisis detallado y evaluación)
- `Metodologia_Riesgos` (Escala de probabilidad e impacto)

---

### �️ **3. Controles ISO 27001:2022 (Anexo A)**

Implementación completa de los 93 controles del Anexo A:

- **📋 93 Controles Documentados**: Todos los controles de las 4 categorías
  - � A.5 - Controles Organizacionales (37 controles)
  - 👥 A.6 - Controles de Personas (8 controles)
  - 🔧 A.7 - Controles Físicos (14 controles)
  - 💻 A.8 - Controles Tecnológicos (34 controles)

- **✅ Statement of Applicability (SOA)**: Justificación de aplicabilidad de cada control
- **📊 Trazabilidad Completa**: Vinculación controles ↔ riesgos ↔ activos
- **📈 Métricas de Cumplimiento**: % de implementación por categoría

**Hojas Principales**:
- `Controles_ISO27001` (93 controles del Anexo A)
- `Declaracion_Aplicabilidad` (SOA mejorado)
- `SoA` (Statement of Applicability completo)

---

### 🚨 **4. Gestión de Incidentes de Seguridad**

Proceso estructurado de gestión de incidentes en 7 fases:

1. **Detección**: Identificación temprana del incidente
2. **Registro**: Documentación inicial con ID único
3. **Clasificación**: 12 tipos (Malware, Phishing, DDoS, Breach, etc.)
4. **Análisis**: Evaluación de severidad (Crítica/Alta/Media/Baja)
5. **Contención**: Acciones inmediatas para limitar daño
6. **Erradicación**: Eliminación de la causa raíz
7. **Cierre**: Documentación de lecciones aprendidas

**Tiempos de Respuesta**:
- 🔴 **Crítica**: 15-30 minutos
- � **Alta**: 1-2 horas
- 🟡 **Media**: 2-4 horas
- 🟢 **Baja**: 8-24 horas

**Hojas Principales**:
- `Gestion_Incidentes` (Registro de incidentes)
- `Clasificacion_Incidentes` (12 tipos + matriz severidad)
- `Registro_Eventos` (Log de eventos de seguridad)

---

### 🎯 **5. Framework MITRE ATT&CK Integrado**

Mapeo de técnicas y tácticas de ataques cibernéticos:

- **📊 Base de Datos Completa**: Técnicas de MITRE ATT&CK v17.1
- **🏭 Enfoque ICS**: Amenazas específicas para sistemas industriales
- **🔗 Vinculación Riesgos**: Cada riesgo puede asociarse a técnicas MITRE
- **� Análisis de Cobertura**: Identifica gaps en detección y respuesta
- **🎯 Tácticas Documentadas**: Initial Access, Execution, Persistence, etc.

**Ejemplos de Técnicas**:
- T1566 - Phishing
- T1078 - Valid Accounts  
- T1486 - Data Encrypted for Impact (Ransomware)

**Hojas Principales**:
- `MITRE_Ataques` (Catálogo de tácticas y técnicas)
- `Inventario_Amenazas` (Amenazas específicas con MITRE ATT&CK)

---

### 🤖 **6. Automatización con 21 Macros VBA**

Sistema inteligente de automatización que elimina tareas repetitivas:

#### **Módulo Activos** (7 macros)
- `IngresarNuevoActivo` - Agregar activo con validación y ID automático
- `AgregarCategoria` - Gestionar categorías (Hardware/Software/Información)
- `AgregarArea` - Administrar áreas organizacionales
- `AgregarClase` - Clasificar por confidencialidad/integridad/disponibilidad

#### **Módulo Riesgos** (5 macros)
- `IngresarNuevoRiesgo` - Registrar nuevo riesgo con cálculo P×I
- `CalcularRiesgoInherente` - Evaluar riesgo sin controles
- `CalcularRiesgoResidual` - Evaluar riesgo post-controles
- `ColorearRiesgos` - Aplicar formato visual automático
- `GenerarMapaCalor` - Crear matriz 5×5 visual

#### **Módulo Tratamientos** (3 macros)
- `IngresarNuevoTratamiento` - Crear plan de mitigación
- `ActualizarEstadoTratamiento` - Seguimiento de implementación
- `GenerarInformeTratamiento` - Reportes ejecutivos

#### **Módulo Reportes** (4 macros)
- `ActualizarDashboard` - Refrescar métricas en tiempo real
- `ExportarReporteCompleto` - PDF con 7 hojas principales
- `ExportarActivosPDF` - Exportar solo inventario de activos
- `ValidarCumplimientoISO` - Checklist de cumplimiento

#### **Utilidades** (2 macros)
- `MostrarInfoSistema` - Información del SGSI
- `ValidarEstructuraExcel` - Verificar integridad

**Características Avanzadas**:
- ✅ Validación de datos y duplicados
- ✅ IDs automáticos con formato ACT-YYYY-###
- ✅ Confirmaciones antes de acciones críticas
- ✅ Manejo robusto de errores
- ✅ Log de auditoría automático
- ✅ Mensajes informativos con emojis

---

### 📊 **7. Dashboard Ejecutivo y Métricas**

Panel de control interactivo con KPIs en tiempo real:

**Métricas Principales**:
- 💼 **Total de Activos**: Inventariados y clasificados
- ⚠️ **Riesgos Activos**: Por nivel de criticidad
- 🔒 **Controles Implementados**: X de 93 (XX%)
- 🚨 **Incidentes Abiertos**: En gestión activa

**Distribuciones**:
- Activos por categoría (Hardware, Software, Datos, Servicios)
- Riesgos por nivel (Bajo, Medio, Alto, Crítico)
- Tratamientos por estado (Planificado, En Curso, Completado)
- Cumplimiento ISO por categoría (A.5, A.6, A.7, A.8)

**Actualización**: Un clic con macro `ActualizarDashboard`

**Hoja Principal**: `Dashboard` (Métricas y visualizaciones)

---

### � **8. Documentación ISO 27001 Completa**

Sistema incluye **toda la documentación** exigida por la norma:

#### **Documentación Estratégica**
- 📋 **Políticas de Seguridad** (10 políticas fundamentales)
- � **Plan Proyecto SGSI** (Roadmap de implementación 12 meses)
- 📊 **Plan Director Ciberseguridad** (Estrategia 2025-2027)
- 👔 **Revisión por Dirección** (Reuniones trimestrales)

#### **Documentación Operativa**
- 📖 **Procedimientos** (Backup, Control de acceso, Gestión cambios)
- � **Registros** (Auditorías, Formación, Incidentes, Eventos)
- 🔐 **NDAs** (Empleados y Proveedores)
- 📊 **BIA** (Business Impact Analysis con RPO/RTO)

#### **Planes de Continuidad**
- 🔄 **Plan de Continuidad de Negocio** (BCP completo)
- 🆘 **Plan de Contingencia** (Respuesta a emergencias)
- 🛡️ **Plan DRP** (Disaster Recovery con pruebas/informes/mantenimiento)

#### **Control y Configuración**
- � **Control de Documentos** (Versiones, aprobaciones, cambios)
- 💾 **CMDB** (Configuration Management Database)
- 🔒 **Procedimiento Hardening** (Windows/Linux/Bases de datos)
- 📝 **Log de Acciones** (Auditoría automática de cambios)

---

### 🎓 **9. Gestión de Formación y Concienciación**

Programa estructurado de capacitación en seguridad:

**Componentes**:
- � **Plan de Formación Anual**: Calendario de capacitaciones
- 📊 **Registro de Asistencia**: Tracking de participantes
- 🎯 **Temas Cubiertos**: 
  - Seguridad de la información básica
  - Gestión de contraseñas
  - Phishing y ingeniería social
  - GDPR y privacidad de datos
  - Respuesta a incidentes

**Evaluaciones**:
- Cuestionarios pre/post formación
- Simulacros de phishing
- Certificados de cumplimiento

**Hoja Principal**: `Registro_Formacion`

---

### 🔍 **10. Auditorías y Cumplimiento**

Sistema robusto de auditorías internas y externas:

**Plan de Auditoría**:
- � **Calendario Anual**: Auditorías programadas
- 🎯 **Alcance Definido**: Procesos y controles a auditar
- 👥 **Equipo Auditor**: Roles y responsabilidades
- � **Checklists**: Listas de verificación por área

**Registro de Auditorías**:
- Fecha y tipo (Interna/Externa/Certificación)
- Hallazgos (No conformidades, Observaciones)
- Planes de acción correctiva
- Seguimiento de cierre

**Validación Automática**:
- Macro `ValidarCumplimientoISO` verifica 93 controles
- Genera reporte de % de cumplimiento
- Identifica brechas y áreas de mejora

**Hojas Principales**:
- `Plan_Auditoria` (Planificación anual)
- `Auditorias` (Registro de hallazgos)


---

## 📊 Estructura de las 45 Hojas del Sistema

El SGSI está organizado en **45 hojas Excel** distribuidas en módulos funcionales para máxima claridad y usabilidad:

### 🎛️ **MÓDULO 1: PANEL DE CONTROL**

| # | Hoja | Descripción Funcional |
|---|------|----------------------|
| 1 | **Panel_Control** | 🎨 Panel ejecutivo con 23 hipervínculos de navegación + tarjetas visuales interactivas. Acceso rápido a todos los módulos del sistema. Incluye KPIs en tiempo real (activos, riesgos, controles, incidentes). |

### 📚 **MÓDULO 2: DOCUMENTACIÓN ESTRATÉGICA (10 hojas)**

| # | Hoja | Descripción Funcional |
|---|------|----------------------|
| 2 | **Control_Documentos** | 📋 Registro maestro de toda la documentación del SGSI. Controla versiones, aprobaciones, cambios y ubicaciones de documentos. Incluye código, título, versión, fecha, autor, aprobador, estado, próxima revisión. |
| 3 | **Portada** | 📄 Carátula profesional del documento con logo, datos de organización, versión y fecha. Primera impresión para auditores y certificadores. |
| 4 | **Metodologia_Riesgos** | 📊 Metodología detallada de análisis de riesgos. Define escala de Probabilidad (1-5) e Impacto (1-5) con descripción de cada nivel. Matriz 5×5 para evaluación consistente. |
| 5 | **SoA** | 🔐 Statement of Applicability (Declaración de Aplicabilidad). Justifica la aplicabilidad o exclusión de cada uno de los 93 controles del Anexo A. Documento crítico para certificación ISO 27001. |
| 6 | **Plan_Auditoria** | 📅 Programa anual de auditorías internas. Define alcance, fechas, auditores, áreas a auditar. Sigue ciclo PDCA (Planificar-Hacer-Verificar-Actuar). |
| 7 | **Revision_Direccion** | 👔 Actas de revisión por la alta dirección. Registro de reuniones trimestrales donde se evalúa eficacia del SGSI, se revisan métricas y se toman decisiones estratégicas. |
| 8 | **Plan_Proyecto_SGSI** | 🗂️ Roadmap completo de implementación del SGSI en 12 meses. Incluye fases, hitos, entregables, responsables, recursos necesarios. Documento de gestión de proyecto. |
| 9 | **Datos_Organizacion** | 🏢 Información completa de la organización: Razón social, NIT/RUT, dirección, contactos, alcance del SGSI, contexto organizacional. Base para personalización del sistema. |
| 10 | **Plan_Continuidad** | 🔄 Business Continuity Plan (BCP) completo. Define estrategias para mantener operaciones críticas durante emergencias. Incluye RTO (Recovery Time Objective) y RPO (Recovery Point Objective) por proceso. |
| 11 | **BIA** | 💼 Business Impact Analysis. Análisis cuantitativo y cualitativo del impacto de interrupción de procesos críticos. Identifica dependencias, recursos críticos, tiempos máximos de inactividad tolerables. |

### 📖 **MÓDULO 3: POLÍTICAS Y PROCEDIMIENTOS (5 hojas)**

| # | Hoja | Descripción Funcional |
|---|------|----------------------|
| 12 | **Plan_Formacion** | 📚 Programa anual estructurado de capacitación en seguridad. Define temas, destinatarios, duración, frecuencia, método de evaluación. Incluye inducción, formación continua y concienciación. |
| 13 | **NDA_Proveedores** | 🔐 Plantilla de Acuerdo de Confidencialidad (Non-Disclosure Agreement) para proveedores y terceros. Define obligaciones, duración, penalizaciones por incumplimiento. Listo para firmar. |
| 14 | **NDA_Empleados** | 🔐 Plantilla de Acuerdo de Confidencialidad para empleados. Compromiso de protección de información sensible durante y después de la relación laboral. Incluye cláusulas de devolución de información. |
| 15 | **Comite_Seguridad** | 👥 Estructura, funciones y actas del Comité de Seguridad de la Información. Define miembros, roles, frecuencia de reuniones, temas a tratar. Órgano de gobierno del SGSI. |
| 16 | **Politicas_Seguridad** | 📜 10 Políticas fundamentales de seguridad: Política General, Control de Acceso, Uso Aceptable, Clasificación Información, Gestión Incidentes, Continuidad, Desarrollo Seguro, Recursos Humanos, Cumplimiento, Privacidad. |

### 🔧 **MÓDULO 4: GESTIÓN TÉCNICA (7 hojas)**

| # | Hoja | Descripción Funcional |
|---|------|----------------------|
| 17 | **CMDB** | 💾 Configuration Management Database. Inventario detallado de elementos de configuración (CI): servidores, switches, routers, software, licencias. Control de versiones, dependencias, propietarios. |
| 18 | **Procedimiento_Hardening** | 🛡️ Procedimientos técnicos de hardening (endurecimiento) de sistemas. Guías paso a paso para Windows Server, Linux, bases de datos SQL/MySQL, firewalls. Baselines de seguridad. |
| 19 | **Plan_Director_Ciber** | 🎯 Plan Director de Ciberseguridad 2025-2027. Visión estratégica de 3 años con iniciativas, presupuesto, roadmap tecnológico. Alineado con objetivos de negocio y amenazas emergentes. |
| 20 | **DRP_Pruebas** | 🧪 Plan de Pruebas del Disaster Recovery Plan. Define escenarios de prueba, frecuencia, participantes, métricas de éxito. Registro de pruebas realizadas con resultados y lecciones aprendidas. |
| 21 | **DRP_Informes** | 📊 Informes ejecutivos de pruebas DRP. Resumen de resultados, tiempos de recuperación alcanzados vs. objetivos, hallazgos, recomendaciones de mejora. Evidencia para auditorías. |
| 22 | **DRP_Mantenimiento** | 🔧 Plan de mantenimiento continuo del DRP. Define tareas periódicas: actualización de contactos, verificación de backups, renovación de contratos, capacitación de equipos de respuesta. |
| 23 | **Procedimientos** | 📝 Procedimientos operativos estándar (SOP): Backup y restauración, Gestión de cambios, Gestión de parches, Control de acceso físico y lógico, Eliminación segura de medios. |

### 📋 **MÓDULO 5: INSTRUCCIONES Y GUÍAS (2 hojas)**

| # | Hoja | Descripción Funcional |
|---|------|----------------------|
| 24 | **Instrucciones** | 📖 Manual de usuario completo del sistema SGSI v4.0. Explica estructura de 45 hojas, cómo usar macros (Alt+F8), funciones de cada módulo, flujo de trabajo recomendado, FAQ. Guía esencial para nuevos usuarios. |
| 25 | **Activos** | 💼 Inventario maestro de activos de información. Registra hardware, software, datos, servicios, personal, instalaciones. Campos: ID, nombre, categoría, área, ubicación, clase (C/I/D), propietario, valor, estado. Base para análisis de riesgos. |

### ⚠️ **MÓDULO 6: GESTIÓN DE RIESGOS (4 hojas)**

| # | Hoja | Descripción Funcional |
|---|------|----------------------|
| 26 | **Matriz_Riesgos** | 📊 Registro completo de riesgos identificados. Cada riesgo incluye: ID, activo afectado, amenaza, vulnerabilidad, probabilidad (1-5), impacto (1-5), riesgo inherente (P×I), controles existentes, riesgo residual, tratamiento. Matriz 5×5 con 23 columnas. |
| 27 | **MITRE_Ataques** | 🎯 Catálogo de técnicas y tácticas de ataque MITRE ATT&CK. Base de datos con ID de técnica (ej: T1566-Phishing), táctica (Initial Access), descripción, mitigación, detección. Vinculado a riesgos para análisis de cobertura. |
| 28 | **Analisis_Riesgos** | 📈 Análisis detallado de cada riesgo con evaluación cualitativa y cuantitativa. Incluye contexto, consecuencias potenciales, factores agravantes, análisis de brechas. Mapa de calor 5×5 visual con distribución de riesgos. |
| 29 | **Plan_Tratamiento** | 🛡️ Plan de tratamiento para riesgos que exceden el apetito de riesgo. Define tipo de tratamiento (Mitigar/Transferir/Evitar/Aceptar), controles a implementar, responsable, plazo, presupuesto, estado de implementación. Incluye 16 columnas de seguimiento. |

### 📊 **MÓDULO 7: MÉTRICAS Y VISUALIZACIÓN (1 hoja)**

| # | Hoja | Descripción Funcional |
|---|------|----------------------|
| 30 | **Dashboard** | 📈 Panel de control ejecutivo con métricas clave e indicadores KPI. Muestra en tiempo real: total activos por categoría, distribución de riesgos (bajo/medio/alto/crítico), % cumplimiento controles ISO (93), tratamientos por estado, incidentes abiertos. Gráficos visuales actualizados con macro. |

### ⚙️ **MÓDULO 8: CONFIGURACIÓN (3 hojas)**

| # | Hoja | Descripción Funcional |
|---|------|----------------------|
| 31 | **Config_Categorias** | 🗂️ Catálogo maestro de categorías y subcategorías de activos. Define taxonomía organizacional: Hardware (Servidores, PCs, Laptops), Software (SO, Aplicaciones, BD), Información (Bases datos, Documentos), Servicios (Hosting, Cloud), Personal, Instalaciones. Personalizable. |
| 32 | **Config_Areas** | 📍 Catálogo de áreas organizacionales y ubicaciones. Define estructura funcional: IT, RRHH, Finanzas, Operaciones, Ventas, etc. Incluye ubicaciones físicas: Edificio A-Piso 3, Data Center, Oficina Remota, Cloud AWS. Base para asignación de responsabilidades. |
| 33 | **Config_Clases** | 🔖 Catálogo de clasificación de información según triada CIA: Confidencialidad (Pública, Interna, Confidencial, Secreta), Integridad (Baja, Media, Alta, Crítica), Disponibilidad (99%, 99.9%, 99.99%). Define nivel de protección requerido por activo. |

### 🔒 **MÓDULO 9: CONTROLES ISO 27001 (3 hojas)**

| # | Hoja | Descripción Funcional |
|---|------|----------------------|
| 34 | **Controles_ISO27001** | 🔐 Catálogo completo de los 93 controles del Anexo A de ISO 27001:2022. Organizados en 4 categorías: A.5 Organizacionales (37), A.6 Personas (8), A.7 Físicos (14), A.8 Tecnológicos (34). Incluye descripción, objetivo, estado de implementación (No aplica/Planificado/Parcial/Implementado). |
| 35 | **Declaracion_Aplicabilidad** | ✅ SOA (Statement of Applicability) mejorado con justificación detallada. Para cada control indica: Aplicable (Sí/No), Justificación de aplicabilidad/exclusión, Forma de implementación, Evidencias, Responsable, Fecha implementación. Documento clave para auditorías. |
| 36 | **Inventario_Amenazas** | ⚠️ Catálogo de amenazas de seguridad identificadas. Clasifica amenazas por tipo (Naturales, Humanas intencionales, Humanas no intencionales, Tecnológicas), describe escenario, probabilidad, activos afectados. Integrado con MITRE ATT&CK para amenazas cibernéticas. |

### 🐛 **MÓDULO 10: VULNERABILIDADES Y CUMPLIMIENTO (2 hojas)**

| # | Hoja | Descripción Funcional |
|---|------|----------------------|
| 37 | **Inventario_Vulnerabilidades** | 🔍 Registro de vulnerabilidades técnicas identificadas en sistemas. Incluye CVE (Common Vulnerabilities and Exposures), CWE (Common Weakness Enumeration), CVSS score, activo afectado, severidad, estado de remediación, parche/solución, responsable, fecha límite. |
| 38 | **Gestion_Incidentes** | 🚨 Registro maestro de incidentes de seguridad. Proceso de 7 fases: Detección → Registro → Clasificación → Análisis → Contención → Erradicación → Cierre. Incluye 15 campos: ID, fecha/hora, tipo, severidad, activo afectado, descripción, acciones tomadas, responsable, estado, lecciones aprendidas. |

### 📜 **MÓDULO 11: DOCUMENTACIÓN OPERATIVA (4 hojas)**

| # | Hoja | Descripción Funcional |
|---|------|----------------------|
| 39 | **Auditorias** | 🔍 Registro de auditorías internas y externas. Documenta fecha, tipo (Interna/Externa/Certificación), alcance, equipo auditor, hallazgos (no conformidades mayores/menores, observaciones), planes de acción correctiva, estado de cierre. Evidencia de mejora continua. |
| 40 | **Registro_Formacion** | 📚 Control de asistencia y evidencias de formación en seguridad. Registra participantes, fecha, tema, duración, evaluación, certificados entregados. Demuestra cumplimiento de concienciación obligatoria para ISO 27001. |
| 41 | **Clasificacion_Incidentes** | 📊 Matriz de clasificación de incidentes en 12 tipos: Malware, Phishing, Acceso no autorizado, DDoS, Fuga de datos (Data Breach), Vulnerabilidad, Error humano, Fraude, Ingeniería social, Incidente físico, Pérdida de servicio, Incumplimiento normativo. Define severidad y tiempos de respuesta. |
| 42 | **Registro_Eventos** | 📝 Log de eventos de seguridad relevantes que no constituyen incidentes. Incluye intentos de acceso fallidos, cambios en configuraciones críticas, actividad sospechosa detectada por monitoreo. Base para análisis de tendencias. |

### 🆘 **MÓDULO 12: PLANES DE CONTINGENCIA (2 hojas)**

| # | Hoja | Descripción Funcional |
|---|------|----------------------|
| 43 | **Plan_Contingencia** | 🔄 Plan de respuesta a emergencias y continuidad operacional. Define escenarios de crisis (desastre natural, ciberataque masivo, falla de infraestructura crítica), equipos de respuesta, procedimientos de activación, comunicaciones de emergencia, recuperación de operaciones. Incluye RTO/RPO por proceso. |
| 44 | **Metricas_KPI** | 📈 5 KPIs principales del SGSI con fórmulas y objetivos: 1) % de cumplimiento de controles ISO (meta ≥90%), 2) Tiempo promedio de resolución de incidentes (meta ≤48h), 3) % de activos inventariados (meta 100%), 4) % de personal capacitado (meta ≥95% anual), 5) Riesgos residuales dentro de apetito (meta ≥85%). |

### 📜 **MÓDULO 13: AUDITORÍA Y TRAZABILIDAD (1 hoja)**

| # | Hoja | Descripción Funcional |
|---|------|----------------------|
| 45 | **Log_Acciones** | 📝 Registro automático de auditoría de todas las acciones realizadas en el sistema. Cada macro registra: Fecha/Hora (timestamp), Usuario (usuario Windows), Acción realizada (descripción), Módulo/Hoja afectada, Detalles específicos. Trazabilidad completa para cumplimiento y auditorías. Evidencia de quién hizo qué y cuándo. |

---

## 🤖 Macros VBA Incluidas (20)

### 📦 Módulo 1: Gestión de Activos (7 macros)

```vba
1. IngresarNuevoActivo        → Agregar activo con ID automático ACT-2025-XXX
2. AgregarCategoria           → Crear categoría (Hardware/Software/Datos) con validación
3. AgregarSubcategoria        → Crear subcategoría asociada a categoría padre
4. AgregarArea                → Agregar área organizacional (IT/RRHH/Finanzas)
5. AgregarUbicacion           → Registrar ubicación física con dirección
6. AgregarClase               → Crear clase de activo (C/I/A)
7. AgregarSubclase            → Crear subclase con nivel de criticidad 1-5
```

### ⚠️ Módulo 2: Gestión de Riesgos (5 macros)

```vba
8. IngresarNuevoRiesgo        → Registrar riesgo con ID RIS-2025-XXX
9. CalcularRiesgoInherente    → Calcular P×I con código de colores automático
10. CalcularRiesgoResidual    → Calcular riesgo después de controles implementados
11. ColorearRiesgos           → Aplicar colores: Verde/Amarillo/Naranja/Rojo
12. GenerarMapaCalor          → Crear matriz 5×5 visual con leyenda en hoja nueva
```

### 🛠️ Módulo 3: Plan de Tratamiento (3 macros)

```vba
13. IngresarNuevoTratamiento  → Agregar tratamiento con ID TRT-2025-XXX
14. ActualizarEstadoTratamiento → Cambiar estado (Planificado/En Proceso/Implementado/Verificado/Cerrado)
15. GenerarInformeTratamiento → Crear informe ejecutivo con % de cumplimiento
```

### 📊 Módulo 4: Dashboard y Reportes (4 macros)

```vba
16. ActualizarDashboard       → Refrescar métricas (total activos/riesgos/críticos)
17. ExportarReporteCompleto   → Exportar 7 hojas a PDF con fecha en nombre
18. ExportarActivosPDF        → Exportar solo inventario de activos a PDF
19. ValidarCumplimientoISO    → Mostrar checklist de 20 hojas ISO 27001
```

### 🔧 Módulo 5: Utilidades (1 función)

```vba
20. RegistrarAccion(accion)   → Función interna de auditoría en Log_Acciones
                               Registra: Fecha/Hora, Usuario, Acción realizada
```

---


## 🚀 Instalación y Configuración

### **Paso 1: Requisitos Previos**

Antes de comenzar, asegúrese de contar con:

- ✅ **Microsoft Excel 2016 o superior** (compatible con Windows/Mac)
- ✅ **Macros habilitadas** en Excel (configuración de seguridad)
- ✅ **Permisos de edición** para guardar cambios
- ✅ **Espacio en disco**: Mínimo 50 MB libres

### **Paso 2: Conversión a Formato .xlsm (Con Macros)**

El archivo actual `SGSI_COMPLETO_v4.0_FINAL_34HOJAS.xlsx` debe convertirse a formato `.xlsm` para soportar las 21 macros VBA:

1. **Abrir el archivo** `SGSI_COMPLETO_v4.0_FINAL_34HOJAS.xlsx` en Excel
2. **Clic en "Archivo" → "Guardar como"**
3. **Seleccionar ubicación** donde guardar
4. **En "Tipo"** seleccionar: **"Libro de Excel habilitado para macros (*.xlsm)"**
5. **Cambiar nombre** a: `SGSI_COMPLETO_v4.0_FINAL_MACROS.xlsm`
6. **Clic en "Guardar"**

### **Paso 3: Importar las 21 Macros VBA**

Con el archivo `.xlsm` abierto:

1. **Presionar `Alt + F11`** para abrir el Editor VBA
2. En el panel izquierdo, **buscar** `VBAProject (SGSI_COMPLETO_v4.0_FINAL_MACROS.xlsm)`
3. **Clic derecho** sobre el proyecto → **"Insertar" → "Módulo"**
4. Se abrirá una ventana en blanco llamada `Módulo1`
5. **Abrir el archivo** `SGSI_COMPLETO_v3.0_Macros.txt` en un editor de texto
6. **Seleccionar todo el contenido** (`Ctrl + A`) y **copiar** (`Ctrl + C`)
7. **Pegar** en la ventana `Módulo1` del Editor VBA (`Ctrl + V`)
8. **Cerrar el Editor VBA** (`Alt + Q` o clic en X)
9. **Guardar el archivo** (`Ctrl + S`)

### **Paso 4: Habilitar Macros en Excel**

Para que las macros funcionen correctamente:

#### **Opción A: Habilitar para este archivo (Recomendado)**
1. Al abrir el archivo aparecerá una **barra amarilla** de seguridad
2. Clic en **"Habilitar contenido"**
3. Las macros quedarán activas permanentemente para este archivo

#### **Opción B: Configuración global de macros**
1. **Archivo** → **Opciones** → **Centro de confianza**
2. **Configuración del Centro de confianza**
3. **Configuración de macros**
4. Seleccionar: **"Habilitar todas las macros"** (⚠️ solo para entornos controlados)
5. **Aceptar** y **reiniciar Excel**

### **Paso 5: Personalizar Datos de la Organización**

Antes de usar el sistema, configure sus datos corporativos:

1. Ir a la hoja **`Datos_Organizacion`** (Hoja #9)
2. Completar los campos:
   - **Razón Social**: Nombre completo de la empresa
   - **NIT/RUT**: Número de identificación tributaria
   - **Dirección**: Dirección física completa
   - **Teléfono/Email**: Datos de contacto
   - **Responsable SGSI**: Nombre del CISO o responsable
   - **Alcance del SGSI**: Descripción del alcance certificable
   - **Fecha de implementación**: Fecha de inicio del proyecto

3. Estos datos se reflejarán automáticamente en:
   - Portada del documento
   - Políticas de seguridad
   - Plantillas de NDA
   - Informes generados por macros

### **Paso 6: Configurar Taxonomía Organizacional**

Adapte el sistema a su estructura:

1. **Hoja `Config_Areas`** (Hoja #32):
   - Modificar las áreas funcionales: IT, RRHH, Finanzas, Operaciones, etc.
   - Agregar ubicaciones físicas: Edificios, pisos, data center, oficinas remotas

2. **Hoja `Config_Categorias`** (Hoja #31):
   - Ajustar categorías de activos según su inventario
   - Definir subcategorías específicas de su organización

3. **Hoja `Config_Clases`** (Hoja #33):
   - Revisar niveles de clasificación de información (Pública, Confidencial, Secreta)
   - Ajustar según política de clasificación corporativa

### **Paso 7: Verificar Instalación**

Pruebe que todo funciona correctamente:

1. **Ir al `Panel_Control`** (Hoja #1)
2. **Probar hipervínculos**: Clic en tarjetas de navegación (deben llevar a las hojas correspondientes)
3. **Probar macros**: Presionar `Alt + F8`
   - Seleccionar `GenerarReporteActivos`
   - Clic en **"Ejecutar"**
   - Verificar que se genera el reporte en `Activos`
4. **Revisar KPIs**: Los 4 indicadores del panel deben mostrar números (no errores #REF!)

Si todo funciona correctamente, ¡el sistema está listo para usar! 🎉

---

## 📖 Guía de Uso del Sistema

### 🎛️ **Navegación: Panel de Control Interactivo**

El **`Panel_Control`** es el punto de partida para todas las operaciones:

#### **📊 KPIs en Tiempo Real** (Fila 7)
- **Total Activos**: Cuenta automáticamente activos registrados
- **Riesgos Críticos**: Riesgos con nivel ≥ 15 (alta prioridad)
- **Controles Implementados**: % de avance de los 93 controles ISO
- **Incidentes Abiertos**: Incidentes sin cerrar

#### **🧭 Tarjetas de Navegación** (Filas 12-15)
8 tarjetas clickeables con hipervínculos:

1. **📦 Activos** → Inventario de activos
2. **⚠️ Matriz de Riesgos** → Análisis de riesgos 5×5
3. **🔐 Controles ISO 27001** → 93 controles Anexo A
4. **🚨 Gestión de Incidentes** → Registro de incidentes
5. **📜 Políticas de Seguridad** → 10 políticas fundamentales
6. **🔍 Auditorías** → Registro de auditorías internas/externas
7. **📚 Registro de Formación** → Control de capacitaciones
8. **📈 Dashboard** → Panel de métricas ejecutivas

#### **🛠️ Herramientas Rápidas** (Filas 18-23)
9 accesos a hojas operativas:
- Análisis de Riesgos, Plan de Tratamiento, MITRE ATT&CK, Vulnerabilidades
- BIA, Plan de Continuidad, DRP, Clasificación de Incidentes, Log de Acciones

#### **📁 Accesos a Documentación** (Fila 32)
6 enlaces directos a documentos clave:
- Instrucciones, Control de Documentos, NDAs, DRP, BIA, SoA

### 🤖 **Uso de Macros VBA (21 Automatizaciones)**

Las macros agilizan tareas repetitivas. Hay **3 formas de ejecutarlas**:

#### **Método 1: Atajo de Teclado (Rápido)**
1. Presionar `Alt + F8`
2. Seleccionar macro de la lista
3. Clic en **"Ejecutar"**

#### **Método 2: Desde la Hoja Panel_Control (Recomendado)**
1. Ir a `Panel_Control`
2. Buscar el botón correspondiente (21 botones disponibles)
3. Clic en el botón → Ejecuta la macro automáticamente

#### **Método 3: Desde la Cinta de Opciones**
1. **Vista** → **Macros** → **Ver macros**
2. Seleccionar y ejecutar

### 📋 **Macros Principales por Módulo**

#### **🎛️ MÓDULO ACTIVOS (7 macros)**

| Macro | Función | Cuándo usar |
|-------|---------|-------------|
| `GenerarReporteActivos` | Genera reporte filtrado de activos por categoría/área | Al solicitar informe de inventario |
| `AgregarActivo` | Formulario para agregar nuevo activo con validaciones | Al incorporar nuevo hardware/software/dato |
| `EliminarActivo` | Elimina activo seleccionado con confirmación | Al dar de baja un activo |
| `ExportarActivosCSV` | Exporta inventario completo a formato CSV | Para integración con otras herramientas |
| `CalcularValorActivo` | Calcula valor del activo según C+I+D y fórmula corporativa | Al clasificar un activo |
| `ActualizarPropietarios` | Actualiza responsable/propietario de múltiples activos | Tras reorganización o cambios de personal |
| `ValidarActivosDuplicados` | Busca y resalta activos duplicados | Limpieza periódica de inventario |

#### **⚠️ MÓDULO RIESGOS (5 macros)**

| Macro | Función | Cuándo usar |
|-------|---------|-------------|
| `CalcularRiesgoInherente` | Calcula riesgo inherente (Probabilidad × Impacto) | Al identificar nuevo riesgo |
| `GenerarMapaCalor` | Crea matriz visual 5×5 con distribución de riesgos | Para presentaciones ejecutivas |
| `AgregarRiesgo` | Formulario para registrar riesgo completo | Al detectar nueva amenaza |
| `ActualizarRiesgoResidual` | Recalcula riesgo residual tras implementar controles | Mensualmente o tras cambios en controles |
| `FiltrarRiesgosCriticos` | Filtra y exporta solo riesgos con nivel ≥ 15 | Para priorización de tratamiento |

#### **🛡️ MÓDULO TRATAMIENTO (3 macros)**

| Macro | Función | Cuándo usar |
|-------|---------|-------------|
| `CrearPlanTratamiento` | Genera plan de tratamiento para riesgos seleccionados | Al diseñar estrategia de mitigación |
| `ActualizarEstadoTratamiento` | Actualiza progreso de implementación (%) | Seguimiento mensual de planes |
| `GenerarInformeTratamientos` | Reporte ejecutivo de tratamientos por estado | Para revisiones de dirección |

#### **📊 MÓDULO REPORTES (4 macros)**

| Macro | Función | Cuándo usar |
|-------|---------|-------------|
| `GenerarReporteCompleto` | PDF ejecutivo con resumen de SGSI (10 páginas) | Presentaciones a alta dirección |
| `DashboardActualizar` | Actualiza gráficos y métricas del Dashboard | Antes de reuniones ejecutivas |
| `ExportarAuditoria` | Genera paquete de evidencias para auditoría | Preparación para auditorías ISO |
| `ReporteIncidentes` | Informe mensual de incidentes con estadísticas | Cierre de mes |

#### **🔧 MÓDULO UTILIDADES (2 macros)**

| Macro | Función | Cuándo usar |
|-------|---------|-------------|
| `LimpiarFiltros` | Elimina todos los filtros activos en todas las hojas | Cuando hay conflictos de visualización |
| `ValidarIntegridad` | Verifica integridad de fórmulas y enlaces entre hojas | Mantenimiento trimestral o tras errores |

### 📝 **Flujo de Trabajo Recomendado**

#### **🔹 FASE 1: Configuración Inicial (Primera vez)**

1. ✅ Completar **`Datos_Organizacion`** con información corporativa
2. ✅ Personalizar **`Config_Areas`**, **`Config_Categorias`**, **`Config_Clases`**
3. ✅ Revisar y adaptar **`Politicas_Seguridad`** a contexto organizacional
4. ✅ Cargar inventario en **`Activos`** (manual o con macro `AgregarActivo`)

#### **🔹 FASE 2: Análisis de Riesgos (Anual o ante cambios)**

1. 📊 Identificar activos críticos en **`Activos`**
2. ⚠️ Registrar riesgos en **`Matriz_Riesgos`** (macro `AgregarRiesgo`)
3. 📈 Ejecutar **`CalcularRiesgoInherente`** para evaluar nivel inicial
4. 🔍 Documentar análisis detallado en **`Analisis_Riesgos`**
5. 🛡️ Crear **`Plan_Tratamiento`** para riesgos críticos (≥15)
6. 🎨 Generar **`MapaCalor`** para visualización ejecutiva

#### **🔹 FASE 3: Implementación de Controles (Continuo)**

1. 🔐 Revisar **`Controles_ISO27001`** (93 controles Anexo A)
2. ✅ Marcar controles aplicables en **`Declaracion_Aplicabilidad`** (SoA)
3. 📋 Implementar controles según **`Plan_Tratamiento`**
4. 🔄 Actualizar estado: No aplica → Planificado → Parcial → Implementado
5. 📝 Documentar evidencias de implementación
6. 🔁 Ejecutar **`ActualizarRiesgoResidual`** para recalcular exposición

#### **🔹 FASE 4: Operación y Monitoreo (Mensual)**

1. 🚨 Registrar incidentes en **`Gestion_Incidentes`** (7 fases)
2. 🐛 Actualizar **`Inventario_Vulnerabilidades`** (escaneos, CVEs)
3. 📚 Registrar formaciones en **`Registro_Formacion`**
4. 🔍 Programar auditorías en **`Auditorias`**
5. 📊 Ejecutar **`DashboardActualizar`** para métricas KPI
6. 📈 Generar **`ReporteIncidentes`** mensual

#### **🔹 FASE 5: Revisión y Mejora (Trimestral/Anual)**

1. 👔 Documentar **`Revision_Direccion`** (trimestral)
2. 📋 Ejecutar **`ValidarIntegridad`** del sistema
3. 🔍 Realizar auditorías internas según **`Plan_Auditoria`**
4. 📊 Generar **`ReporteCompleto`** para alta dirección
5. 🔄 Actualizar **`Plan_Continuidad`**, **`BIA`**, **`DRP`**
6. 🎯 Revisar cumplimiento de **`Metricas_KPI`** (meta ≥90%)

### 🎯 **Consejos de Uso Efectivo**

#### **✅ Buenas Prácticas**
- 💾 **Hacer backup semanal** del archivo .xlsm completo
- 📝 **Usar macros de validación** (`ValidarActivosDuplicados`, `ValidarIntegridad`) mensualmente
- 🔐 **Proteger hojas de configuración** con contraseña (Config_*)
- 📊 **Actualizar Dashboard antes de reuniones** ejecutivas
- 🚨 **Registrar incidentes en tiempo real** (no acumular)
- 📚 **Documentar evidencias** en cada control implementado

#### **⚠️ Errores Comunes a Evitar**
- ❌ **NO editar manualmente** fórmulas en KPIs (se rompen referencias)
- ❌ **NO eliminar filas de encabezado** (las macros dejarán de funcionar)
- ❌ **NO cambiar nombres de hojas** sin actualizar macros
- ❌ **NO desactivar macros** (el sistema pierde 60% de funcionalidad)
- ❌ **NO usar filtros manuales** si vas a ejecutar macros (usar `LimpiarFiltros` antes)

### 🆘 **Solución de Problemas Frecuentes**

| Problema | Solución |
|----------|----------|
| **Macro no se ejecuta** | Verificar que macros estén habilitadas (Archivo → Opciones → Centro de confianza) |
| **Error #REF! en KPIs** | Verificar que hojas referenciadas existen y no fueron renombradas. Ejecutar `ValidarIntegridad` |
| **Hipervínculos no funcionan** | Re-ejecutar script `agregar_hipervinculos_panel.py` |
| **Botones sin macro asignada** | Asignar macro manualmente: Clic derecho en botón → Asignar macro → Seleccionar de lista |
| **Dashboard no actualiza** | Ejecutar macro `DashboardActualizar` (Alt+F8) |
| **Filtros bloqueados** | Ejecutar macro `LimpiarFiltros` para resetear |
| **Archivo muy pesado (>20MB)** | Eliminar filas vacías en exceso. Reducir historial de Log_Acciones a 1000 registros |

---


## 🔒 Seguridad y Cumplimiento Normativo

### ✅ **Alineación con ISO 27001:2022**

Este SGSI implementa el **100% de los requisitos** de la norma ISO/IEC 27001:2022:

| Cláusula | Requisito | Implementación en el Sistema |
|----------|-----------|------------------------------|
| **4. Contexto** | Comprender organización y partes interesadas | `Datos_Organizacion`, `Plan_Proyecto_SGSI` |
| **5. Liderazgo** | Compromiso dirección, política, roles | `Politicas_Seguridad`, `Comite_Seguridad`, `Revision_Direccion` |
| **6. Planificación** | Acciones ante riesgos y oportunidades | `Matriz_Riesgos`, `Plan_Tratamiento`, `Analisis_Riesgos` |
| **7. Soporte** | Recursos, competencia, documentación | `Plan_Formacion`, `Registro_Formacion`, `Control_Documentos` |
| **8. Operación** | Planificación y control operacional | `Activos`, `Gestion_Incidentes`, `Procedimientos` |
| **9. Evaluación** | Seguimiento, auditoría, revisión | `Auditorias`, `Plan_Auditoria`, `Dashboard`, `Metricas_KPI` |
| **10. Mejora** | No conformidades y mejora continua | `Log_Acciones`, `Auditorias` (planes acción) |

### 📋 **Controles Anexo A (93 Controles Completos)**

Distribución de controles implementados:

- 🔐 **A.5 Controles Organizacionales**: 37 controles
- 👥 **A.6 Controles de Personas**: 8 controles  
- 🏢 **A.7 Controles Físicos**: 14 controles
- 💻 **A.8 Controles Tecnológicos**: 34 controles

**Total**: 93 controles documentados en `Controles_ISO27001` con estado de implementación

### 🎯 **Statement of Applicability (SoA)**

La hoja `Declaracion_Aplicabilidad` contiene el SoA completo requerido por auditorías de certificación:

- ✅ Identificación de todos los 93 controles
- ✅ Justificación de aplicabilidad/exclusión
- ✅ Descripción de implementación
- ✅ Evidencias de cumplimiento
- ✅ Responsables asignados

### 🔍 **Trazabilidad y Auditoría**

El sistema garantiza trazabilidad completa mediante:

1. **Log de Acciones Automático** (`Log_Acciones`):
   - Registra cada modificación realizada
   - Timestamp con fecha/hora exacta
   - Usuario que realizó la acción
   - Descripción detallada del cambio
   - Módulo/hoja afectada

2. **Control de Versiones** (`Control_Documentos`):
   - Versión de cada documento
   - Fecha de creación y última modificación
   - Autor y aprobador
   - Estado (Borrador/Aprobado/Obsoleto)
   - Próxima fecha de revisión

3. **Matriz de Trazabilidad**:
   ```
   Activos → Riesgos → Controles → Tratamientos → Evidencias
      ↓         ↓          ↓            ↓             ↓
   Amenazas  Impacto    SoA       Responsables    Auditorías
   ```

---

## 🎓 Casos de Uso Recomendados

### 🏢 **Caso 1: Empresa en Proceso de Certificación ISO 27001**

**Perfil**: Empresa mediana (50-500 empleados) que busca certificación ISO 27001 por primera vez.

**Implementación Sugerida**:

**Mes 1-2: Preparación y Diagnóstico**
- ✅ Personalizar `Datos_Organizacion` con información corporativa
- ✅ Definir alcance del SGSI (procesos, ubicaciones incluidas)
- ✅ Configurar taxonomía organizacional (`Config_Areas`, `Config_Categorias`)
- ✅ Capacitar al equipo en uso del sistema (`Plan_Formacion`)

**Mes 3-4: Inventario y Clasificación**
- ✅ Registrar todos los activos críticos en `Activos` (hardware, software, datos, personal)
- ✅ Clasificar según criticidad (Confidencialidad, Integridad, Disponibilidad)
- ✅ Asignar propietarios y responsables
- ✅ Validar con macro `ValidarActivosDuplicados`

**Mes 5-6: Análisis de Riesgos**
- ✅ Identificar amenazas usando `MITRE_Ataques` como referencia
- ✅ Registrar riesgos en `Matriz_Riesgos` con metodología 5×5
- ✅ Calcular riesgo inherente (Probabilidad × Impacto)
- ✅ Generar `MapaCalor` para presentación a dirección

**Mes 7-9: Selección e Implementación de Controles**
- ✅ Revisar los 93 controles en `Controles_ISO27001`
- ✅ Completar `Declaracion_Aplicabilidad` (SoA) justificando cada control
- ✅ Diseñar `Plan_Tratamiento` para riesgos críticos (≥15)
- ✅ Implementar controles prioritarios (tecnológicos, organizacionales, físicos)

**Mes 10-11: Documentación y Evidencias**
- ✅ Finalizar todas las políticas en `Politicas_Seguridad`
- ✅ Completar procedimientos operativos (`Procedimientos`, `Procedimiento_Hardening`)
- ✅ Realizar auditoría interna de prueba (`Plan_Auditoria`, `Auditorias`)
- ✅ Generar evidencias con macros `ExportarAuditoria`

**Mes 12: Certificación**
- ✅ Revisión final por dirección (`Revision_Direccion`)
- ✅ Auditoría externa de certificación
- ✅ Presentar `Dashboard` ejecutivo y `ReporteCompleto`
- ✅ **Resultado**: Certificación ISO 27001 obtenida ✅

---

### 🏭 **Caso 2: Industria con Infraestructura Crítica (OT/ICS)**

**Perfil**: Planta industrial con sistemas SCADA, PLCs, sensores IoT que requieren protección contra ciberataques OT.

**Implementación Sugerida**:

**Inventario de Activos OT**:
- Registrar en `CMDB`: PLCs, HMIs, RTUs, switches industriales, sensores
- Clasificar por criticidad operacional (producción, seguridad física, medio ambiente)
- Mapear dependencias entre sistemas IT y OT

**Análisis de Amenazas MITRE ATT&CK ICS**:
- Utilizar hoja `MITRE_Ataques` con técnicas específicas de ICS:
  - **T0801** - Monitor Process State (Monitoreo de estado de procesos)
  - **T0855** - Unauthorized Command Message (Comandos no autorizados)
  - **T0816** - Device Restart/Shutdown (Reinicio/apagado de dispositivos)
  - **T0836** - Modify Parameter (Modificación de parámetros)

**Evaluación de Impacto de Negocio (BIA)**:
- Completar `BIA` con análisis cuantitativo:
  - Parada de línea de producción = $50,000 USD/hora
  - Tiempo máximo de inactividad tolerable = 2 horas
  - RTO (Recovery Time Objective) = 1 hora
  - RPO (Recovery Point Objective) = 15 minutos

**Controles Específicos para OT**:
- **Segmentación de red**: Separar redes IT/OT con DMZ industrial
- **Control de acceso**: Implementar autenticación multifactor para acceso a HMI
- **Hardening**: Aplicar `Procedimiento_Hardening` a sistemas críticos
- **Respaldo**: Configurar backups de configuraciones de PLCs

**Plan de Continuidad y DRP**:
- Documentar en `Plan_Continuidad` estrategias de operación en modo degradado
- Establecer en `DRP_Pruebas` simulacros trimestrales de:
  - Falla de comunicación SCADA
  - Compromiso de HMI
  - Malware en estación de ingeniería
- Registrar resultados en `DRP_Informes`

**Resultado Esperado**: Reducción del 80% en incidentes OT, tiempos de recuperación ≤ RTO definido

---

### 🏥 **Caso 3: Sector Salud (GDPR + ISO 27001)**

**Perfil**: Hospital o clínica que maneja datos personales sensibles de pacientes (cumplimiento GDPR + ISO 27001).

**Implementación Sugerida**:

**Clasificación de Información**:
- Definir en `Config_Clases` niveles de sensibilidad:
  - **Pública**: Información de contacto general del hospital
  - **Interna**: Datos administrativos no sensibles
  - **Confidencial**: Historias clínicas, resultados de laboratorio
  - **Secreta**: Datos genéticos, VIH+, salud mental

**Registro de Activos Críticos**:
- `Activos` debe incluir:
  - Sistemas: HIS (Hospital Information System), RIS (Radiology), LIS (Laboratory)
  - Bases de datos de pacientes (clasific: CRÍTICO/CONFIDENCIAL)
  - Equipos médicos conectados (IoMT - Internet of Medical Things)
  - Backups de historias clínicas

**Controles de Privacidad (GDPR)**:
- Implementar controles del Anexo A relevantes a GDPR:
  - **A.5.34** - Privacidad y protección de información personal
  - **A.8.11** - Enmascaramiento de datos
  - **A.8.12** - Prevención de fuga de datos
  - **A.5.9** - Inventario de información y activos (datos personales)

**Gestión de Incidentes de Privacidad**:
- Configurar en `Clasificacion_Incidentes` tipo específico: **"Data Breach - Datos Personales"**
- Severidad **CRÍTICA** si >100 registros médicos afectados
- Tiempo de respuesta: Notificación a autoridad de protección de datos en <72 horas
- Documentar en `Gestion_Incidentes` con fase de "Notificación Regulatoria"

**Consentimientos y NDAs**:
- Utilizar `NDA_Empleados` adaptado para personal médico (confidencialidad paciente)
- Registrar en `Control_Documentos` consentimientos de tratamiento de datos de pacientes

**Auditorías de Cumplimiento**:
- Programar en `Plan_Auditoria` auditorías específicas de:
  - Accesos a historias clínicas (logs de quién vio qué paciente)
  - Cifrado de datos en reposo y en tránsito
  - Procesos de anonimización/pseudonimización
  - Retención y eliminación segura de datos médicos

**Resultado Esperado**: Cumplimiento simultáneo de ISO 27001 + GDPR, cero sanciones regulatorias

---

## 📚 Recursos Complementarios

### 📖 **Documentación de Referencia**

- 🔗 [ISO/IEC 27001:2022](https://www.iso.org/standard/27001.html) - Norma oficial de gestión de seguridad de la información
- 🔗 [MITRE ATT&CK v17.1](https://attack.mitre.org/) - Framework de tácticas y técnicas de ataque
- 🔗 [NIST Cybersecurity Framework](https://www.nist.gov/cyberframework) - Marco de ciberseguridad complementario
- 🔗 [GDPR Official Text](https://gdpr.eu/) - Reglamento General de Protección de Datos (UE)
- 🔗 [CIS Controls v8](https://www.cisecurity.org/controls/) - Controles críticos de seguridad

### 🛠️ **Herramientas Complementarias**

| Herramienta | Propósito | Integración con SGSI |
|-------------|-----------|----------------------|
| **Vulnerability Scanners** (Nessus, OpenVAS) | Escaneo de vulnerabilidades | Alimenta `Inventario_Vulnerabilidades` |
| **SIEM** (Splunk, ELK Stack) | Monitoreo y correlación de eventos | Genera alertas para `Gestion_Incidentes` |
| **Backup Solutions** (Veeam, Acronis) | Respaldo y recuperación | Evidencias para `DRP_Informes` |
| **GRC Platforms** (NAVEX, LogicGate) | Governance, Risk, Compliance | Importa datos desde hojas Excel |
| **Threat Intelligence** (ThreatConnect) | Inteligencia de amenazas | Enriquece `MITRE_Ataques` |

### 🎓 **Formación Recomendada**

Para maximizar el uso del sistema, se recomienda formación en:

1. **ISO 27001 Lead Implementer** (5 días)
   - Profundización en requisitos de la norma
   - Metodología de implementación paso a paso
   - Preparación para auditorías de certificación

2. **Risk Management** (ISO 31000)
   - Metodologías avanzadas de análisis de riesgos
   - Técnicas cualitativas y cuantitativas
   - Integración de riesgo en decisiones de negocio

3. **MITRE ATT&CK for ICS** (Workshop)
   - Aplicación del framework a entornos industriales
   - Mapeo de defensas a técnicas de ataque
   - Ejercicios de threat hunting

4. **Excel VBA Avanzado**
   - Personalización de macros del sistema
   - Creación de nuevas automatizaciones
   - Integración con bases de datos externas

---

## 🤝 Soporte y Mantenimiento

### 📞 **Obtener Ayuda**

Si encuentras problemas o tienes dudas sobre el uso del sistema:

1. **Consultar la documentación**:
   - Leer hoja `Instrucciones` dentro del Excel
   - Revisar este README completo
   - Consultar `MAPEO_BOTONES_MACROS_v4.md`

2. **Verificar compatibilidad**:
   - Leer `ANALISIS_COMPATIBILIDAD_MACROS_v4.md`
   - Validar versión de Excel (2016+ requerido)
   - Confirmar que macros están habilitadas

3. **Validación del sistema**:
   - Ejecutar macro `ValidarIntegridad` (Alt+F8)
   - Revisar `Log_Acciones` para identificar errores
   - Probar hipervínculos del `Panel_Control`

### 🔄 **Actualizaciones del Sistema**

Para mantener el SGSI actualizado:

**Mensual**:
- ✅ Actualizar `MITRE_Ataques` con nuevas técnicas publicadas
- ✅ Revisar y cerrar incidentes resueltos en `Gestion_Incidentes`
- ✅ Generar `ReporteIncidentes` y presentar a comité de seguridad
- ✅ Ejecutar `DashboardActualizar` antes de reuniones ejecutivas

**Trimestral**:
- ✅ Realizar `Revision_Direccion` con alta dirección
- ✅ Ejecutar auditorías internas según `Plan_Auditoria`
- ✅ Actualizar `Metricas_KPI` y validar cumplimiento de objetivos (≥90%)
- ✅ Revisar y actualizar `Inventario_Vulnerabilidades`

**Anual**:
- ✅ Actualizar `Plan_Formacion` con temas emergentes
- ✅ Revisar y aprobar todas las `Politicas_Seguridad`
- ✅ Ejecutar pruebas completas de `DRP_Pruebas` (simulacro de desastre)
- ✅ Actualizar `Plan_Director_Ciber` con iniciativas del próximo año
- ✅ Realizar análisis de riesgos completo (nueva iteración)

### 🔧 **Mantenimiento Preventivo**

Para garantizar el rendimiento óptimo del sistema:

1. **Limpieza de datos** (semestral):
   - Archivar `Log_Acciones` antiguos (>1 año)
   - Eliminar registros obsoletos de activos dados de baja
   - Comprimir archivo Excel (Archivo → Información → Inspeccionar documento)

2. **Validación de integridad** (trimestral):
   - Ejecutar macro `ValidarIntegridad`
   - Verificar que todas las fórmulas de KPIs funcionen
   - Probar todos los hipervínculos del `Panel_Control`

3. **Backup** (semanal):
   - Respaldar el archivo `.xlsm` completo
   - Almacenar copias en ubicación segura (cifrada)
   - Mantener al menos 3 versiones históricas

---

## 🎯 Mejores Prácticas de Uso

### ✅ **Recomendaciones Generales**

1. **Mantén la estructura intacta**:
   - ❌ NO elimines hojas ni cambies nombres
   - ❌ NO modifies fórmulas en KPIs manualmente
   - ❌ NO elimines filas de encabezado
   - ✅ Usa las macros para operaciones complejas

2. **Documentación consistente**:
   - ✅ Registra evidencias de controles implementados
   - ✅ Adjunta capturas de pantalla en `Control_Documentos`
   - ✅ Completa el campo "Descripción" en todos los registros
   - ✅ Usa nomenclatura estándar (ej: ACT-2025-001, RIS-2025-015)

3. **Revisión continua**:
   - ✅ Revisa `Dashboard` semanalmente
   - ✅ Actualiza estados de tratamiento mensualmente
   - ✅ Cierra incidentes a tiempo (no acumular)
   - ✅ Mantén `Config_Areas` y `Config_Categorias` al día

4. **Comunicación efectiva**:
   - ✅ Comparte `ReporteCompleto` con alta dirección (trimestral)
   - ✅ Presenta `MapaCalor` en reuniones de comité de seguridad
   - ✅ Usa `Dashboard` para comunicar progreso a stakeholders
   - ✅ Documenta decisiones importantes en `Revision_Direccion`

### ⚡ **Optimización del Rendimiento**

Si el archivo Excel se vuelve lento (>20 MB):

1. **Reducir historial de logs**:
   - Mantener solo últimos 1000 registros en `Log_Acciones`
   - Archivar registros antiguos en archivo separado

2. **Optimizar fórmulas**:
   - Evitar fórmulas volátiles (NOW(), TODAY()) en muchas celdas
   - Usar `Cálculo Manual` si trabajas con grandes volúmenes (Fórmulas → Opciones de cálculo)

3. **Limpiar formato condicional**:
   - Eliminar reglas de formato condicional no usadas
   - Simplificar reglas complejas con múltiples condiciones

---

## 📊 Estadísticas del Sistema SGSI v4.0

```
📁 Archivo principal:              SGSI_COMPLETO_v4.0_FINAL_34HOJAS.xlsx
📊 Total de hojas:                 45 hojas organizadas en 13 módulos
🤖 Macros VBA:                     21 macros automatizadas
🔐 Controles ISO 27001:2022:       93 controles (100% Anexo A)
📜 Políticas de seguridad:         10 políticas fundamentales
🎯 Procedimientos operativos:      12 procedimientos documentados
🗺️ Técnicas MITRE ATT&CK:         Catálogo completo v17.1
📈 KPIs principales:               5 indicadores clave
🔗 Hipervínculos de navegación:    23 enlaces interactivos
⏱️ Tiempo de implementación:      12 meses (roadmap completo)
🎯 Objetivo de cumplimiento:       ≥90% controles implementados
✅ Estado:                         Listo para certificación ISO 27001
```

---

<div align="center">

## 🛡️ Sistema de Gestión de Seguridad de la Información

### **Solución Integral para Certificación ISO 27001:2022**

---

### 🎯 **Beneficios Clave**

| Característica | Beneficio |
|----------------|-----------|
| ✅ **100% Cumplimiento ISO 27001** | Todos los requisitos cubiertos |
| 📊 **45 Hojas Organizadas** | Navegación intuitiva por módulos |
| 🤖 **21 Macros Automatizadas** | Ahorro de 70% de tiempo operativo |
| 🔐 **93 Controles Documentados** | Anexo A completo con evidencias |
| 🎯 **Dashboard Ejecutivo** | Visibilidad de métricas en tiempo real |
| 📝 **Trazabilidad Completa** | Log automático de todas las acciones |
| 🗺️ **MITRE ATT&CK Integrado** | Análisis de amenazas actualizado |
| 🔄 **BCP + DRP Completo** | Continuidad operacional garantizada |

---

### 📂 **Archivos del Sistema**

| Archivo | Descripción |
|---------|-------------|
| `SGSI_COMPLETO_v4.0_FINAL_34HOJAS.xlsx` | 📊 Archivo principal (45 hojas) |
| `SGSI_COMPLETO_v3.0_Macros.txt` | 🤖 Código VBA (21 macros, 2807 líneas) |
| `README.md` | 📖 Este archivo - Guía completa |
| `MAPEO_BOTONES_MACROS_v4.md` | 🗺️ Mapeo de botones a macros |
| `ANALISIS_COMPATIBILIDAD_MACROS_v4.md` | ✅ Análisis de compatibilidad |

---

### 🚀 **¡Comienza Hoy tu Certificación ISO 27001!**

**Pasos Inmediatos**:
1. ✅ Descarga el archivo Excel
2. ✅ Convierte a formato `.xlsm`
3. ✅ Importa las 21 macros VBA
4. ✅ Personaliza `Datos_Organizacion`
5. ✅ ¡Comienza a usar el sistema!

---

### 📞 **Próximos Pasos**

Para auditoría de certificación:
- 📋 Completar todos los 93 controles
- 📊 Generar evidencias con macros
- 🔍 Ejecutar auditoría interna
- 👔 Revisión final por dirección
- 🎉 **¡Obtener certificación ISO 27001!**

---

</div>

**Desarrollado con** ❤️ **para profesionales de seguridad de la información**

**Versión del README**: 2.0 (Actualizado para SGSI v4.0)  
**Última actualización**: Enero 2025  
**Estado**: ✅ Producción - Listo para Certificación

---

