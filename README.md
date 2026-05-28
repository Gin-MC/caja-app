# 🌟 Plastiluz CashRecon — Sistema de Conciliación y Arqueo de Caja Diario

[![Astro](https://img.shields.io/badge/Framework-Astro%20v4-FF5D01?style=for-the-badge&logo=astro&logoColor=white)](https://astro.build/)
[![TypeScript](https://img.shields.io/badge/Language-TypeScript-3178C6?style=for-the-badge&logo=typescript&logoColor=white)](https://www.typescript.org/)
[![CSS3](https://img.shields.io/badge/Styling-Vanilla%20CSS-1572B6?style=for-the-badge&logo=css3&logoColor=white)](https://www.w3.org/Style/CSS/)
[![SheetJS](https://img.shields.io/badge/Excel-xlsx--js--style-217346?style=for-the-badge&logo=microsoftexcel&logoColor=white)](https://github.com/gitbrent/xlsx-js-style)
[![Vite](https://img.shields.io/badge/Bundler-Vite-646CFF?style=for-the-badge&logo=vite&logoColor=white)](https://vitejs.dev/)

Un sistema profesional, moderno y de alto rendimiento diseñado a medida para **Plastiluz** para agilizar, automatizar y auditar el proceso de conciliación diaria de ventas, cobros bancarios (Yape/BCP), ventas a crédito y arqueo de efectivo físico. 

Este proyecto resuelve el clásico problema de negocio en retail y comercio: **el cruce diario de caja, la gestión de pagos divididos, el control de gastos menores y la generación de reportes financieros oficiales en Excel sin fricciones.**

---

## 🚀 Características Principales

*   **📊 Parser Inteligente de Reportes:** Carga y procesa automáticamente reportes de ventas en formato Excel de forma 100% cliente-servidor (FileReader API) para clasificar instantáneamente las Notas de Venta, Boletas, Facturas y Notas de Crédito.
*   **💳 Reconciliación de Transferencias y Pagos Divididos (BCP/Yape):**
    *   Soporte completo para registrar múltiples transferencias por comprobante o registrar abonos parciales/mixtos (efectivo + transferencia).
    *   Posibilidad de ingresar números de operación, horas específicas y marcar estado de verificación (`SI`/`NO`) en tiempo real.
    *   Filtros y selectores interactivos que permiten cambiar métodos de pago al instante **sin alertas molestas**.
*   **🪙 Inventario y Arqueo de Efectivo Físico:** Panel estructurado para el conteo físico de monedas y billetes peruanos, autocalculando subtotales y acumulados instantáneamente.
*   **🍔 Control Dinámico de Egresos y Gastos de Caja (Conceptos):**
    *   Permite añadir conceptos de gastos menores o ingresos adicionales directamente desde la web (ej. *"Almuerzo del personal"*, *"Pago de flete"*).
    *   Calcula el efectivo teórico esperado restando los egresos y sumando las entradas en tiempo real.
*   **🏷️ Agrupamiento y Ordenamiento de Notas de Crédito (NC):**
    *   Agrupación inteligente en la columna respectiva según la serie del documento: Boletas (`BC01`) y Facturas (`FC01`).
    *   Lógica de ordenamiento dinámico: los comprobantes normales se muestran primero ordenados de forma ascendente, mientras que las Notas de Crédito se posicionan **estrictamente al final de sus columnas** para una visualización contable impecable.
*   **📈 Diseño Altamente Interactivo:** La altura de las tablas se ajusta automáticamente al contenido sin barras de scroll internas molestas, facilitando la visualización en pantallas de supervisión.
*   **📥 Exportación a Excel de Alta Fidelidad (`xlsx-js-style`):**
    *   Genera un reporte profesional de dos hojas (`caja` y `banco`) respetando estrictamente la plantilla original del negocio.
    *   Estructura simétrica de columnas: **`(*) Indicador -> Número de Comprobante -> Monto`**.
    *   Las ventas a crédito muestran la palabra `credito` (en amarillo) en el importe, y el valor numérico va a la columna de indicador.
    *   Crea un **cuadro independiente de Ventas a Crédito** al final del resumen.
    *   Formato estricto de **moneda peruana (`S/. #,##0.00`)** y bordes continuos de cuadrícula gris suave en todas las celdas financieras.
    *   Cálculo de fecha seguro a prueba de desfases de zona horaria (UTC Date serial matching).
    *   **Cero advertencias de archivo dañado:** Todos los tamaños de letra e indicadores XML del libro de Excel se manejan estrictamente con enteros para asegurar compatibilidad total al abrir el archivo.

---

## 🎨 Diseño y Estética Premium

El sistema ha sido estructurado bajo los lineamientos visuales de la identidad corporativa de **Plastiluz**:
*   **Colores de la Marca:** Naranja brillante (`#E67E22`) como distintivo de energía, combinado con un elegante fondo Carbón Oscuro (`#2B2521`) para textos, cabeceras y contrastes premium.
*   **Glassmorphism y Elevaciones:** Tarjetas con sombreados suaves (`box-shadow`), gradientes de bienvenida premium y micro-animaciones en botones interactivos (`hover`, `focus`).
*   **Código de Colores Dinámico por Fila:**
    *   🔵 **Banco BCP / Yape:** Sombreado azul claro elegante (`#EBF5FB`) con textos y selectores en contraste azul.
    *   🟡 **Ventas a Crédito:** Sombreado amarillo claro elegante (`#FEF9E7`) con textos y selectores en contraste ocre.
    *   🔴 **Notas de Crédito:** Sombreado rojo claro elegante (`#FDEDEC`) con textos en contraste guinda.
*   **Modales Estratificados:** Las ventanas emergentes (creación de comprobante y confirmación de eliminación) están ordenadas jerárquicamente con `z-index: 2000` para garantizar que las confirmaciones siempre aparezcan de manera prolija sobre los formularios.

---

## 🛠️ Tecnologías Utilizadas

*   **Astro (v4.x):** Framework web de alto rendimiento centrado en HTML para compilaciones estáticas ultra rápidas.
*   **TypeScript:** Tipado estático para garantizar la robustez en la lógica de estados, cálculos contables y transformaciones de datos.
*   **Vanilla CSS:** Estilos personalizados estructurados con variables globales CSS3 para mantener un control milimétrico sobre el responsive y la paleta de Plastiluz.
*   **xlsx-js-style Fork:** Fork avanzado de SheetJS para posibilitar el diseño de celdas con bordes delgados, fuentes específicas, alineaciones a la izquierda en indicadores y rellenos de color profesionales directamente al exportar desde el navegador.

---

## 📂 Estructura del Proyecto

El proyecto está diseñado de forma modular para facilitar la escalabilidad y el mantenimiento:

```text
caja-app/
├── public/                 # Recursos estáticos
├── src/
│   ├── components/         # Componentes modulares reutilizables
│   │   ├── BankTransfersTable.astro   # Tabla de transferencias BCP y Yape
│   │   ├── CashCounter.astro          # Arqueo físico de monedas y billetes
│   │   ├── ConfirmationModal.astro    # Modal universal de confirmación (Z-Index: 2000)
│   │   ├── ExcelUploader.astro        # Drag & drop de reportes diarios
│   │   ├── Header.astro               # Cabecera con marca Plastiluz
│   │   ├── Logo.astro                 # Logo vectorizado en SVG
│   │   ├── NewDocumentModal.astro     # Formulario de creación de comprobantes (Z-Index: 1000)
│   │   ├── ReconciliationSummary.astro # Dashboard con sumatorias y egresos de caja
│   │   └── SalesSummaryTable.astro    # Tablas auto-adaptables de Ventas diarias
│   ├── js/                 # Motores JavaScript
│   │   ├── excelExporter.js           # Exportador profesional Excel con xlsx-js-style
│   │   └── excelParser.js             # Lector de reportes XLS/XLSX
│   ├── layouts/
│   │   └── Layout.astro               # Contenedor base e inyección de Google Fonts (Outfit)
│   └── pages/
│       └── index.astro                # Consola central de la máquina de estados
├── package.json
└── tsconfig.json
```

---

## 🚀 Instalación y Uso Local

Para ejecutar el reconciliador de caja en tu computadora localmente, sigue estos pasos:

1.  **Clonar el repositorio:**
    ```bash
    git clone https://github.com/tu-usuario/plastiluz-cashrecon.git
    cd plastiluz-cashrecon/caja-app
    ```

2.  **Instalar dependencias:**
    ```bash
    npm install
    ```

3.  **Iniciar el servidor de desarrollo:**
    ```bash
    npm run dev
    ```
    El sistema se abrirá automáticamente en tu navegador en: [http://localhost:4321](http://localhost:4321)

4.  **Compilar para Producción:**
    Si deseas generar los archivos estáticos optimizados y listos para subir a cualquier hosting gratuito (como Vercel, Netlify o GitHub Pages):
    ```bash
    npm run build
    ```
    Los archivos compilados listos se guardarán en la carpeta `dist/`.

---

## 📈 Demostración Contable de Conciliación

El sistema calcula la discrepancia final de caja de la siguiente forma:

$$\text{Efectivo Esperado} = (\text{Ventas Totales} - \text{Banco BCP} - \text{Ventas a Crédito}) + \text{Ajustes de Caja}$$

$$\text{Diferencia (Sobra/Falta)} = \text{Efectivo Físico Contado} - \text{Efectivo Esperado}$$

Este enfoque matemático dinámico previene pérdidas financieras y cuadra al centavo el balance diario de tu negocio de forma automatizada.

---

## 👥 Autor

*   **Gino Cotos** - *Desarrollo de Software y Diseño de Sistemas* - [GitHub](https://github.com/tu-usuario)
