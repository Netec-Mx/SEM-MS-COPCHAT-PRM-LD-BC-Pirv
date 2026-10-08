# Práctica guiada 1. Analizar el portafolio, identificar tendencias y riesgos, y generar conclusiones que se reutilizarán en Word

## Metadatos

| Campo | Valor |
|---|---|
| Duración | **25 min** |
| Complejidad | Media |
| Nivel de Bloom | Aplicar + Analizar |
| Aplicación | Excel con Microsoft 365 Copilot |
| Insumo | `Portfolio_Analisis_[Iniciales].xlsx` |
| Resultado | Hoja `Conclusiones` completada y validada |

## Descripción General

Analizarás un portafolio de seguimiento mensual con Copilot en Excel. El objetivo no es aceptar automáticamente la respuesta de la IA, sino usarla para localizar tendencias, concentraciones, valores atípicos y fórmulas de apoyo, y después comprobar los hallazgos contra la tabla del libro.

> [!IMPORTANT]
> Microsoft documenta que Copilot en Excel puede generar resúmenes, tendencias, valores atípicos, fórmulas, gráficos y tablas dinámicas. También indica que el contenido generado debe revisarse y verificarse. La práctica aplica exactamente ese principio.

## Objetivos de Aprendizaje

- [ ] Consultar una tabla de Excel mediante prompts claros y delimitados.
- [ ] Identificar tendencias, concentraciones y valores atípicos sin convertirlos automáticamente en decisiones.
- [ ] Solicitar una fórmula y validar su lógica antes de conservarla.
- [ ] Registrar de tres a cinco conclusiones trazables para reutilizarlas en Word.

## Prerrequisitos

### Conocimientos requeridos

- Uso básico de tablas, filtros y fórmulas de Excel.
- Concepto de prompt: objetivo + contexto + expectativas + fuente.
- Diferencia entre **hallazgo observado**, **interpretación** y **dato pendiente de confirmar**.

### Acceso requerido

- Microsoft 365 Copilot disponible en Excel.
- Copia de trabajo del archivo [`../recursos/Portfolio_Analisis.xlsx`](../recursos/Portfolio_Analisis.xlsx) guardada en OneDrive o SharePoint como `Portfolio_Analisis_[Iniciales].xlsx`.

## Entorno de Laboratorio

No se requiere una versión específica de Windows ni un build fijo de Office. Es suficiente con que:

- el libro abra correctamente;
- la tabla `Portafolio` sea visible;
- la hoja `Conclusiones` sea editable;
- Copilot esté disponible en Excel.

## Instrucciones Paso a Paso

### Paso 1: Reconocer la fuente y proteger el contexto

**Tiempo:** 3 min  
**Objetivo:** confirmar qué datos pueden utilizarse y dónde se registrarán los resultados.

1. Abre tu copia `Portfolio_Analisis_[Iniciales].xlsx`.
2. Abre la hoja **Portafolio** y haz clic dentro de la tabla.
3. Identifica los encabezados existentes. No cambies sus nombres.
4. Abre la hoja **Conclusiones** y observa las columnas preparadas para registrar los hallazgos.
5. Vuelve a **Portafolio**.

> [!NOTE]
> Las conclusiones deben limitarse a lo que puede comprobarse dentro del archivo. Si un dato no aparece en la tabla, mantenlo como **pendiente de confirmar**.

**Criterio de finalización:** puedes indicar cuál es la tabla fuente y dónde guardarás las conclusiones.

---

### Paso 2: Obtener un mapa inicial de los datos

**Tiempo:** 5 min  
**Objetivo:** comprender la estructura antes de interpretar resultados.

1. Abre Copilot en Excel.
2. Usa **modo Chat** si tu interfaz lo permite, para evitar cambios no solicitados en el libro.
3. Envía:

> **PROMPT 1 — MAPA DE LOS DATOS**
>
> ```text
> Analiza únicamente la tabla Portafolio de este libro.
>
> 1. Resume qué representa cada columna.
> 2. Indica el período cubierto por los datos.
> 3. Identifica valores faltantes o inconsistencias visibles, si existen.
> 4. No inventes columnas, causas ni datos que no estén en la tabla.
>
> Devuelve el resultado en una tabla con: elemento revisado, hallazgo y evidencia observable.
> ```

4. Compara la respuesta con los encabezados y filtros reales.
5. Si Copilot menciona una columna inexistente, corrige con:

```text
Usa únicamente los encabezados visibles de la tabla Portafolio y repite el análisis.
```

**Criterio de finalización:** la respuesta coincide con la estructura real del libro.

---

### Paso 3: Detectar tendencias y concentraciones

**Tiempo:** 5 min  
**Objetivo:** obtener hallazgos que puedan reproducirse con filtros u ordenamientos.

Envía:

> **PROMPT 2 — TENDENCIAS Y CONCENTRACIONES**
>
> ```text
> Usa solo la tabla Portafolio.
>
> Identifica:
> - dos tendencias relevantes entre períodos;
> - una posible concentración por activo, clase o sector;
> - el mayor y el menor rendimiento observado;
> - cualquier valor que parezca atípico.
>
> Para cada punto indica qué columnas utilizaste y qué dato concreto debería verificar yo en Excel.
> No conviertas una observación en recomendación.
> ```

Después:

1. Usa filtros u ordenamiento para comprobar al menos **dos** afirmaciones.
2. Si una cifra no coincide, conserva el valor del libro y descarta la afirmación de Copilot.

> [!TIP]
> Un buen hallazgo debe poder reproducirse sin depender de la redacción de Copilot.

**Criterio de finalización:** tienes al menos dos hallazgos comprobados directamente en la tabla.

---

### Paso 4: Explorar riesgo y valores atípicos

**Tiempo:** 5 min  
**Objetivo:** separar evidencia de interpretación.

Envía:

> **PROMPT 3 — RIESGO Y VALORES ATÍPICOS**
>
> ```text
> Revisa la tabla Portafolio y señala hasta tres elementos que merezcan atención por concentración, rendimiento o nivel de riesgo.
>
> Para cada elemento separa:
> 1. Hecho observable.
> 2. Evidencia en la tabla.
> 3. Interpretación posible.
> 4. Información que faltaría para tomar una decisión.
>
> No supongas causas ni recomendaciones finales.
> ```

Valida cada hecho observable con la tabla.

**Criterio de finalización:** puedes distinguir qué está demostrado y qué todavía es interpretación.

---

### Paso 5: Solicitar una fórmula y revisar su lógica

**Tiempo:** 4 min  
**Objetivo:** usar Copilot como apoyo para fórmulas sin aplicarlas a ciegas.

1. Pide una fórmula que calcule una clasificación simple del rendimiento, por ejemplo positivo, neutro o negativo.
2. Usa:

> **PROMPT 4 — FÓRMULA DE APOYO**
>
> ```text
> Propón una fórmula de Excel para clasificar Rendimiento_pct así:
> - Positivo si es mayor que 0.
> - Neutro si es igual a 0.
> - Negativo si es menor que 0.
>
> Explica la lógica antes de aplicarla y usa el nombre real de la columna existente en la tabla.
> ```

3. Revisa la fórmula propuesta.
4. Si decides aplicarla, crea una columna nueva sin sobrescribir datos existentes.
5. Comprueba manualmente al menos un caso positivo y uno negativo.

**Criterio de finalización:** la fórmula coincide con la regla definida y no modifica datos fuente.

---

### Paso 6: Registrar conclusiones reutilizables

**Tiempo:** 3 min  
**Objetivo:** crear el insumo que utilizará Word.

1. Abre la hoja **Conclusiones**.
2. Registra entre **3 y 5** conclusiones.
3. Completa para cada una:
   - hallazgo;
   - evidencia;
   - método de validación;
   - nivel de confianza;
   - aspecto pendiente, si aplica.
4. Guarda el libro.

> [!WARNING]
> No registres como hecho ninguna afirmación que no hayas podido reproducir en la tabla.

**Criterio de finalización:** la hoja `Conclusiones` contiene de 3 a 5 hallazgos verificables y queda guardada.

## Validación y Pruebas

| # | Criterio | Estado |
|---:|---|:---:|
| 1 | El libro de trabajo conserva la tabla `Portafolio`. | ☐ |
| 2 | Se comprobaron manualmente al menos dos hallazgos. | ☐ |
| 3 | La fórmula solicitada fue revisada antes de conservarse. | ☐ |
| 4 | `Conclusiones` contiene entre 3 y 5 registros. | ☐ |
| 5 | Cada conclusión incluye evidencia o método de validación. | ☐ |
| 6 | No se registraron datos inventados como hechos. | ☐ |

## Solución de Problemas

| Situación | Qué hacer |
|---|---|
| Copilot no aparece | Confirma que abriste el archivo con la cuenta corporativa correcta y que Microsoft 365 Copilot está habilitado. |
| Copilot cita columnas inexistentes | Selecciona la tabla correcta y repite el prompt indicando que use solo los encabezados visibles. |
| Una cifra no coincide | Conserva el valor calculado o visible en Excel; vuelve a pedir el análisis delimitando período y columnas. |
| Copilot modifica el libro cuando solo querías analizar | Deshaz el cambio y usa Chat o una instrucción explícita de no modificar el libro. |

## Limpieza

- Conserva `Portfolio_Analisis_[Iniciales].xlsx` para los Capítulos 3 y 6.
- No elimines la hoja `Conclusiones`.
- Cierra Excel cuando el archivo haya terminado de sincronizarse.

## Resumen

Usaste Copilot en Excel para explorar una tabla, identificar tendencias y riesgos, solicitar una fórmula y **validar** sus resultados. El entregable es la hoja `Conclusiones`, que será el único insumo analítico autorizado para construir el memo en Word.