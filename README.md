*   [Descripción general](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/1-overview)
*   [Arquitectura del sistema](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/1.1-system-architecture)
*   [Estructura de la base de datos](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/1.2-database-structure)
*   [Núcleo de procesamiento de nóminas](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/2-payroll-processing-core)
*   [Cálculo de nómina (Boleta)](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/2.1-payroll-calculation-(boleta))
*   [Gestión de CTS](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/2.2-cts-management)
*   [Sistema de aprovisionamiento](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/2.3-provisioning-system)
*   [Días no laborables y casos especiales](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/2.4-non-worked-days-and-special-cases)
*   [Sistema de informes](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/3-reporting-system)
*   [Generación de informes](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/3.1-report-generation)
*   [T01 Sistema de informes y registro en diario](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/3.2-t01-report-system-and-journaling)
*   [Importación/exportación de datos](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/4-data-importexport)
*   [Sistema de importación de datos](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/4.1-data-import-system)
*   [Exportación electrónica de nóminas (PDT 601)](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/4.2-electronic-payroll-export-(pdt-601))
*   [Integración contable](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/5-accounting-integration)
*   [Generación de asientos contables](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/5.1-accounting-entry-generation)
*   [Integración bancaria](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/5.2-banking-integration)

Descripción general
===================

Archivos fuente relevantes

*   [ActualizadorPlanillaComacsa/SysIntegral.vbp](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/ActualizadorPlanillaComacsa/SysIntegral.vbp)
*   [ActualizadorPlanillaRoda/SysIntegral.vbp](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/ActualizadorPlanillaRoda/SysIntegral.vbp)
*   [Formularios/FrmHorasExtras.frm](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/FrmHorasExtras.frm)
*   [Formularios/FrmPromedios.frm](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/FrmPromedios.frm)
*   [Formularios/MDIplared.frm](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm)
*   [Formularios/Plared00.vbp](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/Plared00.vbp)
*   [VBE9.tmp](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/VBE9.tmp)

El Sistema de Nómina VB6 es una aplicación integral para el procesamiento y la gestión de nóminas, desarrollada en Visual Basic 6.0. Este sistema está diseñado específicamente para que las empresas peruanas gestionen todos los aspectos de la nómina, incluyendo la gestión de empleados, el cálculo de salarios, las deducciones legales, la generación de informes y la integración con los sistemas contables. El software gestiona diversos requisitos específicos de Perú, como la CTS (Compensación por Tiempo de Servicio), las AFP (Administración de Fondos de Pensiones) y el cumplimiento de las normas de la SUNAT (Superintendencia Nacional de Administración Tributaria).

Esta página ofrece una descripción general de la arquitectura del sistema, sus componentes principales y su funcionalidad principal. Para obtener información más detallada sobre cada subsistema, consulte sus respectivas páginas wiki.

Arquitectura del sistema
------------------------

El sistema de nóminas utiliza una arquitectura de tres niveles con capas de presentación, lógica de negocio y acceso a datos. Está desarrollado con Visual Basic 6.0 y se conecta a una base de datos SQL Server para el almacenamiento y la recuperación de datos.

#mermaid-bv5a4h67xro{font-family:ui-sans-serif,-apple-system,system-ui,Segoe UI,Helvetica;font-size:16px;fill:#333}@keyframes edge-animation-frame{from{stroke-dashoffset:0}}@keyframes dash{to{stroke-dashoffset:0}}#mermaid-bv5a4h67xro .edge-thickness-normal{stroke-width:1px}#mermaid-bv5a4h67xro .edge-pattern-solid{stroke-dasharray:0}#mermaid-bv5a4h67xro .marker{fill:#999;stroke:#999}#mermaid-bv5a4h67xro .marker.cross{stroke:#999}#mermaid-bv5a4h67xro p{margin:0}#mermaid-bv5a4h67xro .label{font-family:ui-sans-serif,-apple-system,system-ui,Segoe UI,Helvetica;color:#333}#mermaid-bv5a4h67xro .cluster-label span p{background-color:transparent}#mermaid-bv5a4h67xro span{fill:#333;color:#333}#mermaid-bv5a4h67xro .node rect,#mermaid-bv5a4h67xro .node path{fill:#ffffff;stroke:#dddddd;stroke-width:1px}#mermaid-bv5a4h67xro .node .label{text-align:center}#mermaid-bv5a4h67xro .flowchart-link{stroke:#999;fill:none}#mermaid-bv5a4h67xro .edgeLabel{background-color:#ffffff;text-align:center}#mermaid-bv5a4h67xro .labelBkg{background-color:rgba(255,255,255,0.5)}#mermaid-bv5a4h67xro .cluster rect{fill:#f8f8f8;stroke:#dddddd;stroke-width:1px}#mermaid-bv5a4h67xro .cluster span{color:#444}#mermaid-bv5a4h67xro :root{--mermaid-font-family:"trebuchet ms",verdana,arial,sans-serif}

Data Access Layer

Business Logic Layer

Presentation Layer

User Interface Forms

Crystal Reports

Core Modules

Calculation Engine

Import/Export

Data Access

SQL Server Database

El sistema está diseñado como una aplicación de escritorio de Windows con un contenedor MDI (Interfaz de documentos múltiples) que aloja varios formularios secundarios para diferentes áreas funcionales.

Fuentes:[Formularios/MDIplared.frm1-782](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm#L1-L782) [VBE9.tmp1-42](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/VBE9.tmp#L1-L42) [Formularios/Plared00.vbp1-137](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/Plared00.vbp#L1-L137)

Componentes principales
-----------------------

El sistema de nóminas consta de varios módulos interconectados que manejan diferentes aspectos de la gestión de nóminas:

#mermaid-kkz72kdrxf{font-family:ui-sans-serif,-apple-system,system-ui,Segoe UI,Helvetica;font-size:16px;fill:#333}@keyframes edge-animation-frame{from{stroke-dashoffset:0}}@keyframes dash{to{stroke-dashoffset:0}}#mermaid-kkz72kdrxf .edge-thickness-normal{stroke-width:1px}#mermaid-kkz72kdrxf .edge-pattern-solid{stroke-dasharray:0}#mermaid-kkz72kdrxf .marker{fill:#999;stroke:#999}#mermaid-kkz72kdrxf .marker.cross{stroke:#999}#mermaid-kkz72kdrxf p{margin:0}#mermaid-kkz72kdrxf .label{font-family:ui-sans-serif,-apple-system,system-ui,Segoe UI,Helvetica;color:#333}#mermaid-kkz72kdrxf span{fill:#333;color:#333}#mermaid-kkz72kdrxf .node rect{fill:#ffffff;stroke:#dddddd;stroke-width:1px}#mermaid-kkz72kdrxf .node .label{text-align:center}#mermaid-kkz72kdrxf .flowchart-link{stroke:#999;fill:none}#mermaid-kkz72kdrxf .edgeLabel{background-color:#ffffff;text-align:center}#mermaid-kkz72kdrxf .labelBkg{background-color:rgba(255,255,255,0.5)}#mermaid-kkz72kdrxf :root{--mermaid-font-family:"trebuchet ms",verdana,arial,sans-serif}

MDIplared - Main Application

Personnel Management

Payroll Processing

Reporting System

Accounting Integration

Data Import/Export

Configuration & Parameters

Employee Records

Contract Management

Boleta (Payslip) Generation

CTS Processing

Provision Calculation

AFP Calculation

Payroll Reports

Regulatory Reports

Management Reports

Journal Entry Generation

Cost Center Allocation

Excel Import/Export

PDT 601 Export

Bank File Generation

System Parameters

Calculation Factors

Fuentes:[Formularios/MDIplared.frm243-777](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm#L243-L777) [VBE9.tmp44-186](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/VBE9.tmp#L44-L186)

### Módulos principales del sistema

La funcionalidad del sistema se distribuye en las siguientes áreas principales:

1.  **Maestros (Master Data)** - Gestión de datos base como empleados, empresas y parámetros
2.  **Parámetros (Parameters)** - Configuración de factores de cálculo, tasas de AFP y otras variables
3.  **Movimientos (Transacciones)** - Procesamiento de nóminas, préstamos, anticipos y otras transacciones
4.  **Consultas y Reportes** - Generación de diversos informes y exportaciones de datos
5.  **Contabilidad (Accountability)** - Integración con sistemas contables y generación de asientos contables
6.  **Procesos (Processes)** - Operaciones por lotes y funciones de importación/exportación de datos

Fuentes:[Formularios/MDIplared.frm243-741](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm#L243-L741)

Interfaz de usuario
-------------------

The application uses a Multiple Document Interface (MDI) design with a main container window (`MDIplared.frm`) that hosts various child forms. The main window includes:

*   A menu system for navigating to different functional areas
*   A toolbar with common functions (New, Save, Delete, Search, Print, Process, Exit)
*   A status bar showing current user, server, database, and company information

The menu structure organizes functionality into logical groups that reflect the workflow of payroll processing.

Sources: [Forms/MDIplared.frm4-242](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm#L4-L242)

Key Processes
-------------

### Payroll Processing Workflow

The core payroll process follows this general workflow:

DatabaseSystemUserDatabaseSystemUser#mermaid-kou6f1g4fzf{font-family:ui-sans-serif,-apple-system,system-ui,Segoe UI,Helvetica;font-size:16px;fill:#333}@keyframes edge-animation-frame{from{stroke-dashoffset:0}}@keyframes dash{to{stroke-dashoffset:0}}#mermaid-kou6f1g4fzf .actor{stroke:#cccccc;fill:#ffffff}#mermaid-kou6f1g4fzf text.actor>tspan{fill:#333;stroke:none}#mermaid-kou6f1g4fzf .actor-line{stroke:#cccccc}#mermaid-kou6f1g4fzf .messageLine0{stroke-width:1.5;stroke-dasharray:none;stroke:#999999}#mermaid-kou6f1g4fzf #arrowhead path{fill:#999999;stroke:#999999}#mermaid-kou6f1g4fzf #sequencenumber{fill:#999999}#mermaid-kou6f1g4fzf #crosshead path{fill:#999999;stroke:#999999}#mermaid-kou6f1g4fzf .messageText{fill:#333333;stroke:none}#mermaid-kou6f1g4fzf line{fill:#ffffff;stroke-width:2px}#mermaid-kou6f1g4fzf :root{--mermaid-font-family:"trebuchet ms",verdana,arial,sans-serif}Select payroll periodSelect employeesRetrieve employee dataCalculate payroll itemsDisplay calculation resultsReview and approveSave payroll dataGenerate reportsExport to external systems

The system supports different types of payroll processing:

*   Regular monthly payroll
*   Bi-weekly advances
*   CTS deposits (twice a year)
*   Special bonuses (gratifications)
*   Vacation provisions
*   Year-end closing processes

Sources: [Forms/MDIplared.frm404-446](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm#L404-L446) [Forms/MDIplared.frm478-485](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm#L478-L485) [Forms/MDIplared.frm659-668](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm#L659-L668)

### Data Import/Export Capabilities

The system provides extensive capabilities for importing and exporting data:

#mermaid-5jamcf36y5{font-family:ui-sans-serif,-apple-system,system-ui,Segoe UI,Helvetica;font-size:16px;fill:#333}@keyframes edge-animation-frame{from{stroke-dashoffset:0}}@keyframes dash{to{stroke-dashoffset:0}}#mermaid-5jamcf36y5 .edge-thickness-normal{stroke-width:1px}#mermaid-5jamcf36y5 .edge-pattern-solid{stroke-dasharray:0}#mermaid-5jamcf36y5 .marker{fill:#999;stroke:#999}#mermaid-5jamcf36y5 .marker.cross{stroke:#999}#mermaid-5jamcf36y5 p{margin:0}#mermaid-5jamcf36y5 .label{font-family:ui-sans-serif,-apple-system,system-ui,Segoe UI,Helvetica;color:#333}#mermaid-5jamcf36y5 .cluster-label span p{background-color:transparent}#mermaid-5jamcf36y5 span{fill:#333;color:#333}#mermaid-5jamcf36y5 .node rect{fill:#ffffff;stroke:#dddddd;stroke-width:1px}#mermaid-5jamcf36y5 .node .label{text-align:center}#mermaid-5jamcf36y5 .flowchart-link{stroke:#999;fill:none}#mermaid-5jamcf36y5 .edgeLabel{background-color:#ffffff;text-align:center}#mermaid-5jamcf36y5 .labelBkg{background-color:rgba(255,255,255,0.5)}#mermaid-5jamcf36y5 .cluster rect{fill:#f8f8f8;stroke:#dddddd;stroke-width:1px}#mermaid-5jamcf36y5 .cluster span{color:#444}#mermaid-5jamcf36y5 :root{--mermaid-font-family:"trebuchet ms",verdana,arial,sans-serif}

Data Destinations

Data Sources

Export Functions

Import Functions

Import Excel Payroll Data

Import Time Records

Import Attendance Data

Import Profit Sharing Data

PDT 601 Electronic Payroll

Employee List for SUNAT

Bank Transfer Files

Accounting Entries

Excel Files

Time Recording Systems

Attendance Systems

Tax Authority (SUNAT)

Banks

Accounting System

Management Reports

System

The system can import employee data, attendance records, and other information from Excel files and other sources. It can also export data for regulatory compliance (PDT 601), bank transfers, and accounting integration.

Sources: [Forms/MDIplared.frm614-627](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm#L614-L627) [Forms/MDIplared.frm723-736](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm#L723-L736) [Forms/MDIplared.frm764-775](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm#L764-L775)

Calculation Engine
------------------

The system's calculation engine is responsible for computing various payroll components including:

*   Basic salary and allowances
*   Overtime and bonuses
*   Statutory deductions (AFP, tax, health insurance)
*   Social benefits (CTS, vacations, gratifications)
*   Loans and advances

The calculation logic is primarily contained in the following modules:

*   `Formulas.bas` - Contains mathematical formulas for various calculations
*   `Boleta.bas` - Handles payslip generation logic
*   `ClsCalculaBoleta` - Class for payslip calculation

Sources: [VBE9.tmp61](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/VBE9.tmp#L61-L61) [VBE9.tmp157-158](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/VBE9.tmp#L157-L158)

Reporting System
----------------

The system offers a comprehensive set of reports for different purposes:

1.  **Payroll Reports**
    *   Monthly payroll summaries
    *   Detailed payslips
    *   Remuneration reports
    *   Overtime reports
2.  **Regulatory Reports**
    *   AFP (Pension Fund) reports
    *   SUNAT PDT 601 electronic payroll
    *   CTS deposit certificates
    *   Tax withholding certificates
3.  **Management Reports**
    *   Personnel costs by cost center
    *   Provision reports
    *   Accrued compensation reports
    *   Statistical reports

The system uses Crystal Reports as its reporting engine but also exports data to Excel for additional flexibility.

Sources: [Forms/MDIplared.frm502-608](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm#L502-L608)

Database Integration
--------------------

The system connects to a SQL Server database to store and retrieve all payroll data. The main database tables include:

*   `planillas` - Employee master data
*   `plahistorico` - Payroll history
*   `plaquincena` - Bi-weekly advances
*   `plaremunbase` - Base remuneration
*   `placonstante` - System constants and parameters

For more detailed information about the database structure, please refer to [Database Structure](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/1.2-database-structure).

Sources: [Forms/MDIplared.frm797-828](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm#L797-L828)

Integration Capabilities
------------------------

The payroll system integrates with various external systems:

1.  **Accounting System Integration**
    
    *   Generation of journal entries for payroll transactions
    *   Export to accounting systems via files
    *   Cost center allocation
2.  **Banking Integration**
    
    *   Generation of bank transfer files for salary payments
    *   Support for multiple banks and formats
3.  **Regulatory Reporting**
    
    *   PDT 601 Nómina electrónica para SUNAT
    *   AFP-Net para la información de los fondos de pensiones
4.  **Integración con Excel**
    
    *   Importación y exportación de datos mediante archivos Excel
    *   Generación de informes de Excel

Fuentes:[Formularios/MDIplared.frm738-775](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/MDIplared.frm#L738-L775) [Formularios/FrmHorasExtras.frm247-557](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/FrmHorasExtras.frm#L247-L557) [Formularios/FrmPromedios.frm246-453](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/Forms/FrmPromedios.frm#L246-L453)

Requisitos del sistema
----------------------

La aplicación está construida con Visual Basic 6.0 y requiere:

*   Sistema operativo Windows (compatible con Windows XP y posteriores)
*   Microsoft SQL Server para almacenamiento de bases de datos
*   Tiempo de ejecución de Crystal Reports para la funcionalidad de informes
*   Microsoft Excel para funciones de importación y exportación
*   Conectividad de red adecuada al servidor de base de datos

Fuentes:[VBE9.tmp1-42](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/VBE9.tmp#L1-L42) [ActualizadorPlanillaComacsa/SysIntegral.vbp1-42](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/ActualizadorPlanillaComacsa/SysIntegral.vbp#L1-L42) [ActualizadorPlanillaRoda/SysIntegral.vbp1-42](https://github.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/blob/0d442ace/ActualizadorPlanillaRoda/SysIntegral.vbp#L1-L42)

* * *

Esta descripción general proporciona una comprensión general del sistema de nóminas VB6. Para obtener información detallada sobre componentes específicos, consulte las demás secciones de esta documentación:

*   [Arquitectura del sistema](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/1.1-system-architecture) : arquitectura técnica detallada
*   [Estructura de la base de datos](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/1.2-database-structure) : tablas y relaciones de la base de datos
*   [Núcleo de procesamiento de nóminas](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/2-payroll-processing-core) : detalles sobre el motor de cálculo de nóminas
*   [Sistema de informes](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/3-reporting-system) : información sobre las capacidades de generación de informes
*   [Importación/exportación de datos](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/4-data-importexport) : detalles sobre la funcionalidad de importación y exportación
*   [Integración contable](https://deepwiki.com/VisualFullStackWeb/VB6_PlanillaSueldos_CmsHolcim25/5-accounting-integration) : información sobre la integración del sistema contable

La actualización automática aún no está habilitada

Prueba DeepWiki en tu base de código privada con [Devin](https://deepwiki.com/private-repo)[](https://deepwiki.com/private-repo)

### En esta página

*   [Descripción general](#overview)
*   [Arquitectura del sistema](#system-architecture)
*   [Componentes principales](#core-components)
*   [Módulos principales del sistema](#main-system-modules)
*   [Interfaz de usuario](#user-interface)
*   [Procesos clave](#key-processes)
*   [Flujo de trabajo de procesamiento de nóminas](#payroll-processing-workflow)
*   [Capacidades de importación y exportación de datos](#data-importexport-capabilities)
*   [Motor de cálculo](#calculation-engine)
*   [Sistema de informes](#reporting-system)
*   [Integración de bases de datos](#database-integration)
*   [Capacidades de integración](#integration-capabilities)
*   [Requisitos del sistema](#system-requirements)

Pregúntele a Devin sobre VisualFullStackWeb/VB6\_PlanillaSueldos\_CmsHolcim25

Investigación profunda

(()=>{document.currentScript.remove();processNode(document);function processNode(node){node.querySelectorAll("template\[shadowrootmode\]").forEach(element=>{let shadowRoot = element.parentElement.shadowRoot;if (!shadowRoot) {try {shadowRoot=element.parentElement.attachShadow({mode:element.getAttribute("shadowrootmode"),delegatesFocus:element.getAttribute("shadowrootdelegatesfocus")!=null,clonable:element.getAttribute("shadowrootclonable")!=null,serializable:element.getAttribute("shadowrootserializable")!=null});shadowRoot.innerHTML=element.innerHTML;element.remove()} catch (error) {} if (shadowRoot) {processNode(shadowRoot)}}})}})()
