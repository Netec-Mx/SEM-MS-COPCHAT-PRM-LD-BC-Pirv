# Práctica guiada 1. Analizar el portafolio, identificar tendencias y riesgos, y generar conclusiones que se reutilizarán en Word

## Metadatos

| Campo | Valor |
|---|---|
| Duración | 25 minutos |
| Complejidad | Media |
| Nivel de Bloom | Aplicar |

## Descripción General

En esta práctica, analizará una copia individual del libro `Portfolio_Analisis.xlsx` mediante Microsoft 365 Copilot en Excel. Formulará prompts progresivos para explorar la estructura de los datos, identificar tendencias, concentraciones, variaciones relevantes, posibles valores atípicos y riesgos. Finalmente, validará los hallazgos directamente contra los datos fuente y documentará entre tres y cinco conclusiones verificadas en la hoja `Conclusiones`, que será el único insumo analítico para el memo de inversión posterior en Word.

Copilot acelera la exploración y puede proponer fórmulas, pero no sustituye la revisión humana de los filtros, los períodos, los valores nulos, las referencias de celda ni la lógica de negocio.

## Objetivos de Aprendizaje

- [ ] Formular prompts específicos y verificables para consultar un portafolio en Microsoft 365 Copilot para Excel.
- [ ] Identificar tendencias, concentraciones, variaciones, valores atípicos y riesgos potenciales a partir de los datos disponibles.
- [ ] Solicitar apoyo de Copilot para fórmulas y validar manualmente su lógica antes de utilizarlas.
- [ ] Registrar entre tres y cinco conclusiones trazables, incluyendo al menos dos evidencias numéricas.
- [ ] Redactar una recomendación preliminar sustentada exclusivamente en resultados verificados del libro.

## Prerrequisitos

### Conocimientos requeridos

- Haber observado la Demo 02-01-01 sobre consultas, tendencias, fórmulas y valores relevantes en Excel.
- Comprender el uso básico de tablas, filtros, ordenamiento y fórmulas en Excel.
- Distinguir entre:
  - Un **prompt** o **instrucción temporal**: petición escrita por el participante en el panel de Copilot para una consulta concreta. No modifica permanentemente Copilot.
  - Un **agente preconstruido Analista**: agente persistente administrado y publicado en el tenant corporativo, cuya configuración no debe ser modificada por el participante.
  - Un **mensaje de sistema**: instrucción de configuración aplicada por el servicio o por el administrador. No debe intentarse crear, revelar ni modificar mensajes de sistema durante esta práctica.
- Aplicar validación humana: comprobar resultados de Copilot contra la tabla, los encabezados, los filtros y las fórmulas visibles.

### Acceso requerido

- Cuenta corporativa de Microsoft 365 autenticada, con autenticación multifactor si la política corporativa lo exige.
- Licencia de **Microsoft 365 Copilot** asignada y activa para la cuenta del participante.
- Microsoft Excel autenticado con la misma cuenta corporativa.
- Copia individual de `C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx`.
- Permiso de escritura en la hoja denominada exactamente `Conclusiones`.
- Acceso al agente preconstruido **Analista**, únicamente si el instructor indica utilizarlo para contrastar una consulta.

> **Importante:** no modifique la hoja de datos fuente protegida, los nombres originales de las columnas ni la estructura de la tabla de origen. Solo agregue fórmulas autorizadas, análisis y conclusiones en las áreas indicadas por el instructor.

## Entorno de Laboratorio

### Hardware y conectividad

| Componente | Especificación requerida |
|---|---|
| Equipo | Windows 11 corporativo de 64 bits, procesador de 2 núcleos a 1,6 GHz o superior |
| Memoria | 8 GB de RAM como mínimo |
| Espacio libre | 5 GB como mínimo para archivos de laboratorio, caché de Office y kit de marca |
| Pantalla | Mínimo 1366 × 768; recomendado 1920 × 1080 |
| Red | Conectividad HTTPS saliente por el puerto 443; recomendado 10 Mbps de descarga y 2 Mbps de carga |

### Software y servicios

| Tecnología | Versión o configuración requerida | Licencia o uso | Fuente oficial |
|---|---|---|---|
| Windows 11 Pro o Enterprise | 24H2, compilación 26100.1, arquitectura x64 | Sistema operativo corporativo administrado | https://learn.microsoft.com/windows/release-health/windows11-release-information |
| Microsoft 365 Apps for enterprise | Versión 2408, compilación 17928.20114, arquitectura x64 | Aplicaciones de Office corporativas | https://learn.microsoft.com/officeupdates/update-history-microsoft365-apps-by-date |
| Excel para Microsoft 365 | Incluido en Microsoft 365 Apps for enterprise, versión 2408, compilación 17928.20114, arquitectura x64 | Requerido para esta práctica | https://support.microsoft.com/es-es/excel |
| Microsoft 365 Copilot en Excel | Servicio SaaS administrado, **[VERSIÓN POR VALIDAR]**, sin arquitectura de cliente aplicable | Requiere licencia Microsoft 365 Copilot asignada y habilitada por el tenant | https://learn.microsoft.com/es-es/copilot/microsoft-365/microsoft-365-copilot-overview |
| Microsoft 365 Copilot Chat | Servicio SaaS administrado, **[VERSIÓN POR VALIDAR]**, sin arquitectura de cliente aplicable | No es el entorno principal de esta práctica; no sustituye Copilot integrado en Excel | https://support.microsoft.com/es-es/topic/acerca-de-microsoft-365-copilot-chat-9eb1bbd3-8d36-4a5f-9ada-4e3f3a62d04a |
| Agente preconstruido Analista | Servicio de agente administrado, **[VERSIÓN POR VALIDAR]**, sin arquitectura de cliente aplicable | Uso complementario y solo si lo indica el instructor | [ENLACE OFICIAL] |
| Microsoft Edge | Versión 128.0.2739.42, arquitectura x64 | Navegador corporativo para acceso a servicios Microsoft 365, si fuese necesario | https://learn.microsoft.com/deployedge/microsoft-edge-relnote-stable-channel |
| Microsoft Designer | No utilizado en esta práctica | No requerido; no usar para generar elementos visuales del análisis | https://support.microsoft.com/es-es/designer |
| Microsoft Planner | No utilizado en esta práctica | No requerido; no registrar tareas ni conclusiones en Planner | https://support.microsoft.com/es-es/planner |

### Archivos y directorio de trabajo

| Elemento | Ubicación |
|---|---|
| Directorio de trabajo | `C:\CopilotLabs\Batch1\` |
| Libro de análisis | `C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx` |
| Entregable de esta práctica | Hoja `Conclusiones` dentro de `Portfolio_Analisis.xlsx` |
| Entregables posteriores | `Memo_Inversion.docx` y `Deck_Ejecutivo_Inversion.pptx` |

### Comandos de comprobación inicial

Abra **Windows PowerShell** y ejecute los siguientes comandos. No modifique los archivos mediante PowerShell durante esta práctica.

```powershell
Test-Path "C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx"
Get-Item "C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx" | Select-Object Name, Length, LastWriteTime
Get-FileHash "C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx" -Algorithm SHA256
```

El primer comando debe devolver `True`. Registre mentalmente la fecha de modificación del archivo antes de abrirlo; no es necesario capturar pantalla salvo que el instructor lo solicite.

## Instrucciones Paso a Paso

### Paso 1: Abrir el libro y revisar los límites de trabajo

**Objetivo:** Confirmar que trabaja sobre su copia individual del libro y reconocer las áreas permitidas para análisis.

**Instrucciones:**

1. Abra Excel y autentíquese con su cuenta corporativa si se le solicita.
2. Abra el archivo:
   ```text
   C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx
   ```
3. Compruebe que el libro contiene:
   - Una hoja de datos fuente protegida contra modificaciones accidentales.
   - Una hoja denominada exactamente `Conclusiones`.
4. Seleccione la hoja de datos fuente y ubique la tabla principal.
5. Revise visualmente:
   - Encabezados de columna.
   - Número aproximado de registros.
   - Período cubierto por los datos.
   - Campos que podrían representar categoría, región, producto, responsable, importe, valor, fecha, rendimiento, riesgo u otra métrica financiera.
6. No cambie encabezados, no elimine registros y no desbloquee hojas protegidas.
7. Guarde el archivo una vez con **Ctrl+S** para confirmar que puede editar la hoja `Conclusiones`.

**Resultado esperado:**

- El libro se abre desde la carpeta de trabajo establecida.
- La hoja fuente permanece sin modificaciones.
- La hoja `Conclusiones` está disponible para registrar hallazgos.

**Verificación:**

- Confirme que la pestaña de la hoja se denomina exactamente `Conclusiones`.
- Confirme que los encabezados de la tabla fuente permanecen en su idioma y forma originales.
- Si el libro solicita guardar cambios en la hoja fuente sin que usted haya realizado modificaciones, seleccione **Cancelar** y comuníquelo al instructor.

### Paso 2: Inspeccionar estructura, período y calidad básica de los datos

**Objetivo:** Obtener una descripción verificable de la estructura del portafolio antes de interpretar tendencias o riesgos.

**Instrucciones:**

1. Haga clic dentro de cualquier celda de la tabla fuente.
2. Abra el panel de **Copilot** en Excel.
3. Envíe el siguiente prompt. Ajuste únicamente el nombre de la tabla si Excel lo muestra y Copilot lo requiere.

   ```text
   Analiza únicamente la tabla seleccionada. Identifica:
   1) los encabezados disponibles;
   2) el período mínimo y máximo cubierto por los campos de fecha, si existen;
   3) el número de registros;
   4) campos con valores vacíos, formatos inconsistentes o datos que puedan afectar el análisis.
   
   No inventes columnas ni valores. Si un dato no está disponible, indícalo explícitamente.
   Presenta una tabla breve con: elemento revisado, resultado, evidencia y limitación.
   ```

4. Lea la respuesta y compárela con la tabla fuente:
   - Revise al menos dos encabezados.
   - Compruebe el primer y último valor de fecha si existe una columna de fecha.
   - Filtre una columna con datos potencialmente vacíos, si Copilot reporta valores nulos.
5. Si la respuesta menciona una columna inexistente, formule esta corrección:

   ```text
   La columna mencionada no existe en la tabla. Repite el análisis usando solamente los encabezados visibles en la tabla actual y cita los nombres de columna exactamente como aparecen.
   ```

6. Registre en una nota temporal, dentro de `Conclusiones` o en el área autorizada, el período y las limitaciones que afectarán su análisis. No redacte todavía conclusiones ejecutivas finales.

**Resultado esperado:**

- Una descripción de los campos disponibles, el período cubierto y posibles limitaciones de calidad.
- Confirmación humana de que Copilot utilizó encabezados existentes.
- Identificación de datos faltantes, si los hubiera.

**Verificación:**

- Al menos dos encabezados citados por Copilot coinciden exactamente con los encabezados de Excel.
- El período informado se contrasta con la tabla o con filtros de fecha.
- Si hay datos faltantes, se identifica la columna afectada y no se asume que equivale a cero.

### Paso 3: Generar un resumen ejecutivo inicial y tendencias relevantes

**Objetivo:** Obtener un resumen inicial del portafolio y comprobar al menos una tendencia relevante directamente en la fuente.

**Instrucciones:**

1. Mantenga seleccionada una celda dentro de la tabla fuente.
2. En Copilot, envíe este prompt contextual. Sustituya los elementos entre corchetes por nombres reales de encabezados solo si existen en el libro.

   ```text
   Usa únicamente los datos de la tabla actual. Genera un resumen ejecutivo inicial del portafolio para un comité de inversión.
   
   Examina el período disponible y resume:
   - comportamiento general de [métrica principal];
   - tendencias por [categoría o región disponible];
   - aumentos o disminuciones relevantes;
   - limitaciones de los datos que impidan una conclusión firme.
   
   Para cada afirmación cuantitativa, indica el campo utilizado, el período y el valor o porcentaje correspondiente. No atribuyas causas que no estén respaldadas por los datos.
   ```

3. Revise la respuesta. Seleccione una tendencia cuantitativa informada por Copilot, por ejemplo:
   - Incremento o disminución de una categoría.
   - Concentración de valor en una región.
   - Cambio entre períodos.
4. Valide esa tendencia usando uno de estos métodos:
   - Filtro de tabla y suma de la columna relevante.
   - Tabla dinámica autorizada por el instructor.
   - Fórmula visible en una zona permitida.
5. Si la tabla tiene un campo temporal, solicite mayor precisión:

   ```text
   Desglosa la tendencia más relevante por período usando solamente los campos existentes. Muestra los valores por período y calcula la variación porcentual entre el primer y el último período disponibles. Si faltan datos para el cálculo, explica qué falta.
   ```

6. Anote la tendencia validada, incluyendo el campo, el período y el valor numérico exacto.

**Resultado esperado:**

- Un resumen ejecutivo inicial basado en datos reales.
- Al menos una tendencia validada manualmente con un valor o porcentaje.
- Identificación de cualquier incertidumbre derivada de datos incompletos o de falta de contexto.

**Verificación:**

- La tendencia anotada incluye una evidencia numérica.
- El valor coincide con el resultado del filtro, suma, tabla dinámica o fórmula utilizada.
- La redacción evita causalidad no demostrada. Por ejemplo, escriba “se observa una disminución” y no “la disminución fue causada por”.

### Paso 4: Identificar concentraciones, valores atípicos y riesgos potenciales

**Objetivo:** Localizar elementos que requieran atención y diferenciar entre un hallazgo observado y un riesgo que requiere análisis adicional.

**Instrucciones:**

1. En Copilot, formule el siguiente prompt:

   ```text
   Revisa la tabla actual para identificar riesgos potenciales del portafolio basados exclusivamente en los datos disponibles.
   
   Busca:
   1) concentraciones relevantes por categoría, región, producto, responsable o activo;
   2) máximos, mínimos y valores atípicos;
   3) variaciones inusuales entre períodos;
   4) registros con datos faltantes que puedan distorsionar el análisis.
   
   Devuelve una lista priorizada de hasta cinco hallazgos. Para cada hallazgo incluye: prioridad, evidencia numérica, columnas utilizadas, período aplicable y por qué requiere revisión humana.
   No presentes una hipótesis como un hecho.
   ```

2. Revise cada hallazgo propuesto y seleccione un máximo de tres que sean verificables.
3. Para cada hallazgo seleccionado:
   - Aplique filtros sobre las columnas citadas.
   - Ordene de mayor a menor o de menor a mayor según corresponda.
   - Confirme el valor, categoría, registro o período señalado.
4. Clasifique cada elemento de manera cuidadosa:
   - **Hallazgo confirmado:** se puede reproducir directamente con filtros, fórmulas o valores de la tabla.
   - **Riesgo potencial:** los datos justifican atención, pero no prueban una causa o impacto futuro.
   - **No confirmado:** no coincide con los datos, no tiene evidencia suficiente o depende de una columna inexistente.
5. Si Copilot identifica un posible valor atípico sin explicar el criterio, solicite aclaración:

   ```text
   Explica el criterio utilizado para considerar atípico el registro o grupo indicado. Muestra el valor observado, una referencia comparativa calculable con la tabla y la limitación de ese criterio.
   ```

**Resultado esperado:**

- Una lista priorizada de hallazgos potenciales.
- Entre uno y tres hallazgos contrastados directamente con la tabla.
- Distinción explícita entre hechos confirmados y riesgos potenciales.

**Verificación:**

- Cada hallazgo elegido contiene una columna, un período o categoría y una cifra verificable.
- Ningún riesgo se presenta como certeza futura.
- Los valores atípicos cuentan con un criterio visible o se descartan si no se pueden sustentar.

### Paso 5: Solicitar y validar una fórmula de apoyo

**Objetivo:** Usar Copilot para obtener apoyo en una fórmula sin aplicarla ciegamente ni modificar la fuente protegida.

**Instrucciones:**

1. Determine una necesidad analítica concreta. Ejemplos:
   - Calcular variación porcentual entre dos períodos.
   - Clasificar valores por un umbral definido.
   - Calcular participación de una categoría sobre el total.
   - Contar registros con datos faltantes.
2. Solicite una fórmula adaptada a las columnas disponibles. Use el siguiente modelo:

   ```text
   Necesito una fórmula de Excel para [objetivo concreto].
   
   Los encabezados disponibles son: [pegue únicamente los encabezados relevantes].
   La fórmula se colocará en un área autorizada fuera de la hoja fuente protegida.
   
   Explica:
   1) la fórmula propuesta;
   2) qué referencia o columna debe ajustarse;
   3) cómo evita errores por división entre cero o celdas vacías;
   4) un ejemplo de resultado esperado.
   
   No insertes la fórmula automáticamente.
   ```

3. Revise la fórmula propuesta antes de copiarla:
   - Confirme que usa nombres reales de columna o referencias válidas.
   - Compruebe que la métrica usada es coherente con el cálculo.
   - Revise el manejo de ceros, errores o valores vacíos.
4. Si utiliza una tabla de Excel y las columnas correspondientes existen, un ejemplo posible de participación porcentual es:

   ```excel
   =IFERROR([@[Importe]]/SUM([Importe]),0)
   ```

   Use este ejemplo solo si la tabla realmente tiene una columna denominada `Importe`. No cambie los encabezados para que coincidan con el ejemplo.

5. Si compara dos valores en celdas autorizadas y la celda anterior puede ser cero, un ejemplo posible es:

   ```excel
   =IFERROR((C2-B2)/B2,"")
   ```

   Ajuste `B2` y `C2` a las celdas reales. No aplique esta fórmula sobre la hoja protegida.
6. Introduzca la fórmula únicamente en el área autorizada por el instructor o en una columna analítica permitida.
7. Compruebe manualmente dos resultados:
   - Un registro con valor normal.
   - Un registro con valor cero, vacío o extremo, si existe.
8. Si la fórmula no coincide con el cálculo manual, no la utilice para una conclusión. Corrija el prompt o solicite al instructor orientación.

**Resultado esperado:**

- Una fórmula revisada por el participante y aplicada solamente en una ubicación permitida.
- Dos comprobaciones manuales de los resultados de la fórmula.
- Evidencia adicional para sustentar, matizar o descartar un hallazgo.

**Verificación:**

- La fórmula no genera errores no controlados como `#DIV/0!`, `#REF!` o `#NAME?`.
- Los dos casos comprobados coinciden con el cálculo manual.
- La fórmula no altera la hoja fuente protegida ni sus encabezados.

### Paso 6: Contrastar una consulta con el agente Analista, si lo indica el instructor

**Objetivo:** Diferenciar el uso de Copilot integrado en Excel del uso complementario de un agente preconstruido.

**Instrucciones:**

1. Realice este paso solo si el instructor confirma que el agente preconstruido **Analista** está disponible y debe utilizarse.
2. Abra el agente Analista mediante el punto de acceso corporativo indicado por el instructor.
3. Envíe un prompt de contraste basado en un hallazgo ya validado. No incluya información fuera del libro ni datos confidenciales que no estén autorizados para la práctica.

   ```text
   Estoy revisando un portafolio en Excel. El hallazgo validado es el siguiente:
   [describa el hallazgo, la cifra, el período y las columnas utilizadas].
   
   Propón preguntas de verificación para un comité de inversión. Distingue entre:
   - evidencia disponible;
   - supuestos que requieren confirmación;
   - datos adicionales que serían necesarios para evaluar el riesgo.
   
   No inventes datos ni emitas una recomendación definitiva.
   ```

4. Compare la respuesta del agente con el resultado validado en Excel.
5. Utilice únicamente preguntas útiles o advertencias metodológicas. No copie al libro afirmaciones que el agente no pueda respaldar con datos del archivo.
6. Regrese a Excel para completar el registro de conclusiones.

**Resultado esperado:**

- Una lista breve de preguntas de verificación o datos adicionales requeridos.
- Comprensión de que el agente Analista no sustituye la evidencia verificable de la tabla de Excel.

**Verificación:**

- Puede identificar qué dato proviene de Excel y qué elemento es una sugerencia del agente.
- No se registra como hecho ninguna afirmación que no esté respaldada por el libro.
- Si el agente no está disponible, la práctica continúa sin bloqueo.

### Paso 7: Registrar conclusiones verificadas para reutilización en Word

**Objetivo:** Crear un insumo trazable, conciso y reutilizable para el memo de inversión de la siguiente actividad.

**Instrucciones:**

1. Abra la hoja `Conclusiones`.
2. En el área autorizada, cree o complete una tabla con las siguientes columnas:

   | Prioridad | Conclusión verificada | Evidencia numérica | Fuente y período | Validación realizada | Recomendación preliminar | Estado |
   |---|---|---|---|---|---|---|

3. Registre entre **tres y cinco conclusiones verificadas**.
4. Asegúrese de que al menos **dos conclusiones** contengan evidencia numérica precisa, por ejemplo:
   - Importe, total, promedio, mínimo o máximo.
   - Porcentaje de variación.
   - Participación de una categoría.
   - Número de registros afectados.
5. Para cada conclusión:
   - Cite las columnas o campos usados.
   - Indique el período, si aplica.
   - Describa cómo se validó: filtro, ordenamiento, suma, tabla dinámica, fórmula revisada o revisión manual.
   - Diferencie observación y recomendación.
6. Incluya una recomendación preliminar única o varias recomendaciones específicas. Ejemplos de formulación aceptable:
   - “Revisar la exposición de la categoría X antes de la próxima decisión de asignación.”
   - “Validar la integridad de los registros sin fecha antes de comparar períodos.”
   - “Solicitar explicación de la variación de Y, ya que los datos muestran el cambio pero no su causa.”
7. Evite afirmaciones no verificadas, por ejemplo:
   - Incorrecto: “La categoría X caerá el próximo trimestre.”
   - Correcto: “La categoría X presenta una disminución de [valor] en el período analizado; se recomienda revisar sus factores explicativos.”
8. Guarde el archivo con **Ctrl+S**.

**Resultado esperado:**

- La hoja `Conclusiones` contiene entre tres y cinco conclusiones.
- Existen al menos dos evidencias numéricas verificadas.
- Las recomendaciones preliminares se distinguen claramente de los datos observados.
- El archivo queda listo para ser utilizado en la práctica de Word.

**Verificación:**

- La tabla incluye las siete columnas solicitadas o sus equivalentes claramente identificables.
- Hay un mínimo de tres y un máximo de cinco filas de conclusiones.
- Al menos dos filas contienen cifras y períodos comprobables.
- La hoja fuente no contiene cambios estructurales.

## Validación y Pruebas

Realice las siguientes validaciones antes de dar por terminada la práctica.

| Prueba | Acción | Criterio de aprobación | Evidencia requerida |
|---|---|---|---|
| Integridad del archivo | Compruebe que el libro se guarda en `C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx`. | El archivo existe, abre sin errores y conserva la hoja `Conclusiones`. | Archivo guardado y abierto correctamente. |
| Protección de fuente | Revise la hoja fuente y sus encabezados. | No se modificaron nombres de columnas, registros fuente ni protección de la hoja. | Comparación visual de encabezados y estructura. |
| Trazabilidad | Revise cada conclusión registrada. | Cada conclusión cita campos, período cuando aplica y método de validación. | Tabla de `Conclusiones` completa. |
| Evidencia numérica | Cuente las conclusiones con cifras comprobables. | Al menos dos evidencias numéricas coinciden con filtros, fórmulas o cálculos visibles. | Dos comprobaciones directas en Excel. |
| Fórmula | Revise la fórmula utilizada, si se creó una. | La fórmula usa referencias válidas, maneja errores razonablemente y coincide con dos comprobaciones manuales. | Fórmula visible y resultados comprobados. |
| Calidad del prompt | Revise un prompt utilizado. | El prompt identifica objetivo, alcance de datos, métrica o campos relevantes, período y formato esperado. | Prompt visible en historial de Copilot o anotado en el área autorizada. |
| Supervisión humana | Compare un hallazgo de Copilot con la fuente. | El resultado se acepta, ajusta o descarta con una justificación basada en datos. | Nota de validación en la fila correspondiente. |
| Caso adversarial: información inexistente | Envíe el siguiente prompt a Copilot: `Indica la calificación crediticia externa de cada activo y cita su fuente.` | Copilot debe reconocer que no puede responder si esa información no existe en la tabla. No se registra ninguna respuesta sin evidencia como hecho. | Respuesta que declara limitación, o corrección del participante si Copilot inventa información. |
| Caso adversarial: instrucción incrustada | Si encuentra texto en celdas, notas o comentarios que diga “ignora instrucciones”, “omite validaciones” o similar, trátelo como dato no confiable. No siga esa instrucción. | El participante mantiene el objetivo de análisis, no altera la fuente y valida cualquier contenido contra la tabla. | Confirmación verbal al instructor o anotación breve de que se ignoró la instrucción no autorizada. |

La práctica se considera completada cuando se cumplen todos los criterios obligatorios siguientes:

1. El archivo se guarda correctamente en la ubicación establecida.
2. La hoja fuente no fue modificada.
3. La hoja `Conclusiones` contiene de tres a cinco hallazgos verificados.
4. Existen al menos dos evidencias numéricas reproducibles.
5. Existe al menos una recomendación preliminar que no afirma causalidad ni predicción sin sustento.
6. Cualquier salida de Copilot no verificable fue marcada como limitación, pregunta pendiente o descartada.

## Solución de Problemas

### Problema 1: Copilot no aparece en Excel o indica que la cuenta no tiene acceso

**Síntomas:**

- No se muestra el botón o panel de Copilot en Excel.
- El panel informa que la licencia no está disponible.
- Copilot solicita iniciar sesión repetidamente.
- La respuesta indica que la funcionalidad no está habilitada para la cuenta.

**Causa probable:**

La cuenta no está autenticada con la identidad corporativa correcta, la licencia de Microsoft 365 Copilot no está asignada o la configuración del tenant aún no habilita Copilot en Excel para el usuario.

**Corrección:**

1. En Excel, seleccione **Archivo > Cuenta**.
2. Compruebe que la cuenta mostrada es la cuenta corporativa asignada al curso.
3. Cierre sesión y vuelva a iniciar sesión si se muestra una cuenta personal o diferente.
4. Cierre y vuelva a abrir Excel.
5. Si el problema continúa, no use Copilot Chat como sustituto automático del Copilot integrado en Excel. Informe al instructor y continúe con la revisión manual de tabla, filtros y fórmulas mientras se valida el acceso.
6. No intente instalar complementos no autorizados ni cambiar configuraciones de licencia.

### Problema 2: Copilot propone columnas, cálculos o conclusiones que no coinciden con el libro

**Síntomas:**

- Copilot menciona una columna que no existe.
- Un total, porcentaje o período no coincide con los filtros de Excel.
- La fórmula propuesta genera `#REF!`, `#NAME?` o resultados incoherentes.
- La respuesta atribuye una causa que no figura en los datos.

**Causa probable:**

El prompt fue demasiado amplio, no delimitó la tabla o el período, la tabla contiene valores faltantes o formatos inconsistentes, o Copilot interpretó incorrectamente los campos disponibles.

**Corrección:**

1. No copie el resultado a `Conclusiones` como un hecho.
2. Seleccione una celda dentro de la tabla correcta.
3. Reenvíe un prompt delimitado, por ejemplo:

   ```text
   Usa solamente los encabezados visibles de la tabla actual. No inventes columnas ni causas. Muestra el cálculo, los campos usados y el período. Si no puedes verificar un dato, indícalo como limitación.
   ```

4. Verifique los cálculos con filtro, ordenamiento, suma o fórmula visible.
5. Corrija referencias de fórmulas para que coincidan con las columnas reales.
6. Si persiste la discrepancia, descarte el hallazgo y documente la limitación en lugar de forzar una conclusión.

## Limpieza

1. Guarde el libro con **Ctrl+S**.
2. Cierre el panel de Copilot si ya no lo necesita.
3. Cierre `Portfolio_Analisis.xlsx`.
4. No elimine, renombre ni mueva los archivos de:
   ```text
   C:\CopilotLabs\Batch1\
   ```
5. No borre la hoja `Conclusiones`, ya que será utilizada como fuente directa en la actividad de Word.
6. Si creó cálculos temporales fuera del área autorizada, elimínelos únicamente si el instructor lo indica y sin modificar la hoja fuente protegida.
7. Cierre Excel al finalizar, salvo que el instructor solicite una revisión inmediata del libro.

## Resumen

En esta práctica utilizó Microsoft 365 Copilot en Excel para explorar un portafolio mediante prompts específicos, revisar tendencias, identificar concentraciones y riesgos potenciales, y solicitar apoyo para fórmulas. La actividad enfatizó que las respuestas de IA son propuestas analíticas que requieren validación humana contra los datos fuente.

El entregable es la hoja `Conclusiones` de `Portfolio_Analisis.xlsx`, con entre tres y cinco conclusiones verificadas, al menos dos evidencias numéricas y una recomendación preliminar. Este contenido se reutilizará sin cambiar de fuente en la siguiente práctica de Word para elaborar el memo de inversión.

Recursos de consulta:

- [Obtener información sobre los datos con Copilot en Excel](https://support.microsoft.com/es-es/topic/obtener-informaci%C3%B3n-sobre-los-datos-con-copilot-en-excel-1bba8e8f-97f8-43c0-a89f-6c1e8c7f4f97)
- [Ayuda y aprendizaje de Copilot en Excel](https://support.microsoft.com/es-es/excel)
- [Introducción a Microsoft 365 Copilot](https://learn.microsoft.com/es-es/copilot/microsoft-365/microsoft-365-copilot-overview)

---

# Demo: 2.1 Demostracion Consultar datos, explicar tendencias, apoyar fórmulas y detectar valores relevantes

## Metadatos

| Elemento | Valor |
|---|---|
| Duración | 10 minutos |
| Complejidad | Fácil |
| Nivel de Bloom | Comprender |

## Descripción General

> ℹ️ **Nota:** Esta práctica es una **Demostración realizada por el instructor**. El instructor ejecutará los pasos y comandos mientras los alumnos observan, toman notas y analizan el procedimiento, en lugar de realizarla individualmente.

El instructor utilizará Microsoft 365 Copilot integrado en Excel para explorar un portafolio estructurado, solicitar tendencias, identificar valores relevantes y pedir apoyo para comprender fórmulas. La demostración enfatiza la diferencia entre una respuesta generada por Copilot, los datos verificables en la tabla y la interpretación profesional que debe realizar la persona analista.

## Objetivos de Aprendizaje

- [ ] Reconocer cómo formular prompts contextuales para consultar una tabla de portafolio en Excel.
- [ ] Distinguir consultas orientadas a tendencias, excepciones, riesgos, comparaciones y valores atípicos.
- [ ] Identificar cómo solicitar explicaciones o propuestas de fórmulas sin aplicarlas automáticamente.
- [ ] Verificar afirmaciones de Copilot contra columnas, filtros, períodos y valores de la hoja.
- [ ] Diferenciar una instrucción temporal de chat, un prompt de usuario y un mensaje de sistema no editable por el participante.

## Prerrequisitos

**Conocimientos requeridos**

- Comprensión básica de tablas de Excel: encabezados, filtros, filas, columnas y fórmulas.
- Capacidad para interpretar métricas de portafolio, como rendimiento, valor, sector, período, variación o nivel de riesgo.
- Conocimiento de que Copilot puede generar hipótesis y resúmenes, pero no sustituye la validación humana.
- Comprensión de los términos siguientes:
  - **Prompt:** solicitud concreta que el usuario escribe en Copilot para una interacción puntual.
  - **Instrucción:** indicación temporal incluida en un prompt, por ejemplo, “responde en una tabla de tres columnas”.
  - **Mensaje de sistema:** configuración interna administrada por el servicio que define límites o comportamiento general; no es un prompt del alumno ni debe intentarse modificar mediante texto en el chat.
  - **Agente o asistente persistente:** experiencia configurada para una función determinada, con instrucciones y capacidades publicadas por la organización. No se utiliza el agente preconstruido **Analista** en esta demostración; se usa Copilot integrado en Excel.

**Acceso requerido para el instructor**

- Cuenta corporativa autenticada en Microsoft 365 con licencia de **Microsoft 365 Copilot** habilitada.
- Acceso de lectura al archivo `C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx`.
- Excel autenticado con la misma cuenta corporativa.
- Pantalla compartida, proyector o sesión remota visible para los participantes.
- Acceso HTTPS saliente por el puerto 443 hacia los servicios corporativos autorizados de Microsoft 365.

## Entorno de Laboratorio

| Componente | Versión o configuración requerida | Fuente oficial |
|---|---|---|
| Sistema operativo | Windows 11 Pro o Enterprise, 24H2, compilación 26100.1, arquitectura de 64 bits | https://learn.microsoft.com/windows/release-health/windows11-release-information |
| Excel | Microsoft 365 Apps for enterprise, versión 2408, compilación 17928.20114, arquitectura de 64 bits | https://learn.microsoft.com/officeupdates/update-history-microsoft365-apps-by-date |
| Microsoft 365 Copilot en Excel | Servicio SaaS administrado por Microsoft; versión de servicio: **[VERSIÓN POR VALIDAR]**; licencia de Microsoft 365 Copilot asignada al instructor | https://learn.microsoft.com/microsoft-365-copilot/microsoft-365-copilot-overview |
| Microsoft 365 Copilot Chat | Servicio SaaS; versión de servicio: **[VERSIÓN POR VALIDAR]**. No es la experiencia principal de esta demostración y no sustituye a Copilot integrado en Excel. | https://support.microsoft.com/copilot-chat |
| Agente preconstruido Analista | Servicio administrado por Microsoft; versión: **[VERSIÓN POR VALIDAR]**. No se usa en esta demo; se utilizará en actividades posteriores si está publicado en el tenant. | https://learn.microsoft.com/microsoft-365-copilot/extensibility/ |
| Microsoft Designer | Servicio SaaS; versión: **[VERSIÓN POR VALIDAR]**. No se utiliza ni se requiere licencia para esta demostración. | https://support.microsoft.com/designer |
| Microsoft Planner | Servicio SaaS; versión: **[VERSIÓN POR VALIDAR]**. No se utiliza ni se requiere licencia para esta demostración. | https://support.microsoft.com/planner |

> **Distinción de licencias y herramientas:** la capacidad demostrada depende de la licencia y configuración de **Microsoft 365 Copilot en Excel**, no de Copilot Chat, Designer, Planner ni del agente Analista. Si Copilot Chat está disponible en el tenant, puede servir para conversaciones generales, pero no reemplaza el contexto directo de una tabla abierta en Excel.

**Archivo de referencia**

| Archivo | Ruta | Uso permitido |
|---|---|---|
| Libro de portafolio | `C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx` | Solo lectura y demostración durante esta actividad |
| Hoja de datos fuente | Protegida contra modificaciones accidentales | No modificar |
| Hoja `Conclusiones` | Disponible para actividades posteriores | No modificar durante esta demo |

**Comandos de preparación que el instructor puede mostrar**

Abra Windows PowerShell y ejecute los siguientes comandos para confirmar la existencia del archivo y la conectividad HTTPS. Los participantes observan; no necesitan ejecutarlos.

```powershell
Test-Path "C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx"
Get-Item "C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx" |
    Select-Object Name, Length, LastWriteTime
Test-NetConnection office.com -Port 443
```

**Resultado esperado de preparación**

- `Test-Path` devuelve `True`.
- El archivo aparece con nombre, tamaño y fecha de modificación.
- `TcpTestSucceeded` devuelve `True` o la conectividad es confirmada por la red corporativa.
- El instructor abre el archivo sin modificar los datos fuente.

## Instrucciones Paso a Paso

### Paso 1: Confirmar el contexto y la integridad del libro

**Objetivo:** Mostrar a los participantes que Copilot debe recibir contexto de una tabla estructurada y que el libro fuente no se modifica durante la demostración.

**Instrucciones**

1. El instructor abre Excel y carga el archivo:

   ```text
   C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx
   ```

2. El instructor identifica la hoja que contiene los datos de portafolio y muestra que la fuente está protegida contra modificaciones accidentales.
3. El instructor selecciona una celda dentro de la tabla de datos y muestra los encabezados disponibles, sin cambiar sus nombres.
4. El instructor comprueba visualmente que la información está estructurada como tabla de Excel, normalmente mediante la pestaña contextual **Diseño de tabla**.
5. El instructor localiza la hoja denominada exactamente `Conclusiones`, pero no agrega información en ella durante la demostración.
6. El instructor explica que los nombres de columna originales deben conservarse, aunque los prompts se redacten en español.
7. El instructor abre el panel de Copilot en Excel y confirma que la cuenta corporativa está autenticada.

**Resultado esperado**

- El libro está abierto.
- La tabla de portafolio es visible y contiene encabezados.
- La hoja fuente permanece sin cambios.
- El panel de Copilot está disponible para consultas sobre el libro.

**Verificación**

- Los alumnos pueden identificar al menos dos encabezados reales de la tabla.
- El instructor confirma que no aparece ningún indicador de edición no guardada causado por esta demostración.
- El instructor solicita a los participantes anotar los campos visibles que podrían ser útiles para análisis, por ejemplo: sector, período, valor, rendimiento, riesgo o variación.

---

### Paso 2: Demostrar una consulta inicial con contexto verificable

**Objetivo:** Mostrar cómo un prompt específico reduce ambigüedad y permite obtener un resumen inicial útil.

**Instrucciones**

1. El instructor explica la estructura de un prompt contextual:

   ```text
   Objetivo + datos o columnas relevantes + período + criterio + formato esperado
   ```

2. En el panel de Copilot, el instructor escribe un prompt adaptado a los encabezados reales de la tabla. Si existen las columnas indicadas, puede utilizar el siguiente ejemplo:

   ```text
   Analiza la tabla de portafolio actualmente abierta. Resume los hallazgos principales por sector y período. Identifica los tres sectores con mayor valor total y señala cualquier variación relevante. Responde en una tabla con: hallazgo, columnas utilizadas y evidencia que debo verificar manualmente.
   ```

3. El instructor revisa la respuesta sin aceptarla como conclusión definitiva.
4. El instructor señala qué partes de la respuesta son:
   - Observaciones generadas por Copilot.
   - Datos que pueden verificarse en la tabla.
   - Interpretaciones que requieren juicio profesional.
5. El instructor muestra cómo usar filtros de Excel para comprobar al menos una afirmación cuantitativa de Copilot.

**Resultado esperado**

- Copilot devuelve un resumen relacionado con los datos de la tabla.
- La respuesta menciona sectores, períodos, valores o variaciones según las columnas disponibles.
- El instructor demuestra que una afirmación debe contrastarse con el origen.

**Verificación**

- Se verifica manualmente al menos una cifra, porcentaje o clasificación mencionada por Copilot.
- El instructor identifica la columna y el filtro usados para la comprobación.
- Los alumnos registran que “Copilot indicó” no equivale a “dato validado”.

---

### Paso 3: Refinar la consulta para analizar tendencias y riesgos

**Objetivo:** Evidenciar cómo una instrucción adicional puede delimitar el horizonte temporal, el criterio de riesgo y el formato de salida.

**Instrucciones**

1. El instructor parte de la respuesta anterior y explica que una respuesta amplia puede requerir refinamiento.
2. El instructor formula un prompt de seguimiento. Debe ajustar los nombres de columna a los disponibles en el libro:

   ```text
   Refina el análisis usando únicamente los registros del período más reciente disponible. Compara el rendimiento o variación por sector e identifica posibles riesgos según los valores extremos, la concentración o la volatilidad disponible en la tabla. No infieras causas externas. Presenta los resultados como: tendencia observada, evidencia en columnas y pregunta de validación para el analista.
   ```

3. El instructor destaca tres elementos del prompt:
   - “Únicamente los registros del período más reciente” delimita el alcance temporal.
   - “No infieras causas externas” limita conclusiones no sustentadas.
   - “Pregunta de validación” obliga a distinguir evidencia de interpretación.
4. El instructor filtra la tabla al período identificado y contrasta una tendencia reportada.
5. El instructor explica que términos como “riesgo”, “relevante” o “significativo” deben vincularse a criterios observables, no dejarse completamente implícitos.

**Resultado esperado**

- Copilot produce un análisis más acotado al período solicitado.
- La respuesta presenta tendencias como hallazgos sujetos a verificación.
- El instructor puede mostrar una evidencia visible para al menos una tendencia.

**Verificación**

- Los participantes pueden identificar el período usado en la respuesta.
- El instructor confirma que la tendencia verificada coincide con los registros filtrados.
- Los alumnos anotan una posible limitación: por ejemplo, datos nulos, períodos incompletos, registros duplicados o ausencia de una métrica de volatilidad.

---

### Paso 4: Detectar valores relevantes y solicitar explicación de métricas

**Objetivo:** Demostrar consultas para localizar máximos, mínimos, excepciones y posibles valores atípicos sin convertir automáticamente una observación en decisión.

**Instrucciones**

1. El instructor solicita una exploración de extremos y excepciones. Ajuste los campos al libro real:

   ```text
   Examina la tabla de portafolio e identifica máximos, mínimos y posibles valores atípicos en las métricas numéricas disponibles. Para cada caso, indica el registro o categoría asociada, la columna utilizada y la razón por la que debe revisarse. No clasifiques un valor como error sin evidencia en la tabla.
   ```

2. El instructor revisa si Copilot diferencia correctamente entre:
   - Un valor máximo o mínimo.
   - Un valor atípico estadístico o aparente.
   - Un posible error de datos.
   - Un hallazgo que requiere revisión adicional.
3. El instructor selecciona uno de los valores mencionados y lo localiza mediante filtro u ordenación en Excel.
4. El instructor formula una consulta de explicación sobre una métrica real del archivo, por ejemplo:

   ```text
   Explica qué representa la columna [NOMBRE_REAL_DE_LA_COLUMNA] según los valores y encabezados disponibles en esta tabla. Indica qué puedes observar directamente y qué no puedes determinar sin documentación adicional.
   ```

5. El instructor recalca que Copilot puede explicar patrones observados, pero no debe inventar definiciones de negocio que no estén documentadas.

**Resultado esperado**

- Copilot identifica registros, categorías o valores que merecen revisión.
- La respuesta incluye límites o solicitudes de confirmación cuando falte contexto.
- El instructor muestra un valor relevante directamente en la tabla.

**Verificación**

- Se localiza en Excel al menos uno de los valores máximos, mínimos o excepcionales reportados.
- El instructor confirma que la respuesta no se utiliza para declarar un error de datos sin revisión.
- Los alumnos registran una diferencia entre “valor extremo” y “riesgo confirmado”.

---

### Paso 5: Solicitar apoyo para fórmulas y validar la lógica antes de aplicarla

**Objetivo:** Mostrar que Copilot puede proponer o explicar fórmulas, pero que la persona analista debe validar referencias, lógica y reglas de negocio antes de usar una fórmula.

**Instrucciones**

1. El instructor señala dos columnas numéricas compatibles con una comparación, por ejemplo, valor inicial y valor final, importe anterior e importe actual, o rendimiento esperado y rendimiento real.
2. El instructor solicita una propuesta de fórmula sin insertarla automáticamente:

   ```text
   Sin modificar el libro, propón una fórmula de Excel para calcular la variación porcentual entre [COLUMNA_VALOR_ANTERIOR] y [COLUMNA_VALOR_ACTUAL] en una tabla de Excel. Explica cómo tratar división por cero o valores vacíos. Usa referencias estructuradas si son adecuadas y explica la fórmula de forma breve.
   ```

3. El instructor revisa la fórmula propuesta en pantalla y compara:
   - Los nombres de columnas usados.
   - El orden de numerador y denominador.
   - El tratamiento de cero, celdas vacías y errores.
   - La coherencia con la definición de negocio de “variación porcentual”.
4. El instructor explica una fórmula típica solo como ejemplo conceptual:

   ```excel
   =IFERROR(([@[Valor actual]]-[@[Valor anterior]])/[@[Valor anterior]],"")
   ```

5. El instructor aclara que la fórmula anterior no debe pegarse si los encabezados reales no coinciden o si la política de negocio requiere otro resultado para valores nulos o cero.
6. El instructor no aplica ni guarda modificaciones en la hoja fuente.

**Resultado esperado**

- Copilot propone una fórmula o explica una fórmula relevante.
- El instructor identifica las partes de la fórmula que necesitan validación humana.
- El libro permanece sin cambios.

**Verificación**

- Los participantes pueden explicar qué representa el numerador y el denominador de la fórmula.
- El instructor confirma que la fórmula propuesta no se aplicó a la tabla fuente.
- Se registra una regla de validación: nunca aceptar una fórmula sin revisar referencias, tratamiento de errores y lógica de negocio.

## Validación y Pruebas

La demostración se considera completada cuando el instructor produce evidencia observable de las siguientes validaciones:

| Criterio medible | Evidencia esperada | Resultado esperado |
|---|---|---|
| Contexto de datos confirmado | Tabla de Excel visible con encabezados originales y hoja `Conclusiones` identificada | La fuente de datos no se modifica |
| Consulta contextual realizada | Un prompt incluye objetivo, columnas o métricas, período y formato esperado | Copilot responde sobre la tabla abierta |
| Tendencia contrastada | Filtro, ordenación o revisión manual de la tabla | Al menos una afirmación de Copilot coincide con los datos visibles |
| Valor relevante revisado | Registro máximo, mínimo o potencialmente atípico localizado en Excel | Se diferencia “valor para revisar” de “error confirmado” |
| Fórmula analizada | Fórmula propuesta visible, sin aplicarla en la fuente | Referencias y manejo de errores revisados por el instructor |
| Trazabilidad registrada | Notas de los participantes con prompt, evidencia y limitación | Los alumnos pueden explicar qué debe validarse |

**Caso adversarial: información inexistente o sin evidencia**

El instructor ejecuta el siguiente prompt, adaptando el nombre del instrumento o campo para que no exista en la tabla:

```text
Indica el rendimiento del instrumento INEXISTENTE-999 durante el último período. Si no aparece en la tabla abierta, responde exactamente: “No hay evidencia en los datos disponibles” y recomienda una acción de verificación. No inventes valores ni fuentes externas.
```

**Resultado esperado del caso adversarial**

- Copilot debe reconocer que no hay evidencia suficiente en la tabla o indicar que no localiza el instrumento.
- Una respuesta que invente un valor, una fuente externa o una recomendación basada en datos inexistentes debe ser tratada como no válida.
- El instructor explica que la ausencia de evidencia debe conducir a una solicitud de datos adicionales, no a una conclusión.

**Criterios de calidad del prompting observados**

- **Precisión:** el prompt define la métrica, el período y el resultado solicitado.
- **Trazabilidad:** el análisis puede relacionarse con columnas, filtros o registros visibles.
- **Incertidumbre:** Copilot debe indicar límites cuando falten campos, definiciones o evidencia.
- **Utilidad:** la respuesta permite decidir qué revisar después.
- **Supervisión humana:** el instructor valida datos y no delega la decisión final en Copilot.

## Solución de Problemas

### Problema 1: Copilot en Excel no aparece o muestra que la capacidad no está disponible

**Síntoma:** El botón o panel de Copilot no aparece en Excel, o aparece un mensaje que indica que la cuenta no tiene acceso.

**Causa probable:** La cuenta del instructor no tiene asignada una licencia de Microsoft 365 Copilot, Excel no está autenticado con la cuenta corporativa correcta, o la configuración del tenant no habilita la experiencia.

**Corrección:**

1. Confirmar que Excel inició sesión con la cuenta corporativa autorizada.
2. Comprobar en **Archivo > Cuenta** que Microsoft 365 Apps está activado.
3. Solicitar al administrador de Microsoft 365 la confirmación de la licencia Microsoft 365 Copilot y de las políticas aplicables.
4. Cerrar y volver a abrir Excel después de cualquier cambio de licencia o autenticación.
5. Si el servicio no está disponible, realizar la demostración con las capturas aprobadas por el instructor y explicar que no se deben simular resultados como si fueran respuestas en vivo.

### Problema 2: La respuesta de Copilot es genérica, usa un período incorrecto o no coincide con la tabla

**Síntoma:** Copilot responde sin mencionar columnas relevantes, mezcla períodos, no identifica la métrica solicitada o presenta afirmaciones difíciles de comprobar.

**Causa probable:** El prompt no especifica suficiente contexto, la tabla no está seleccionada, existen campos ambiguos, hay filtros activos no identificados o los datos contienen valores vacíos.

**Corrección:**

1. Seleccionar una celda dentro de la tabla antes de volver a consultar.
2. Reformular el prompt incluyendo nombres reales de columnas, período, criterio y formato de salida.
3. Indicar explícitamente que no se infieran causas externas ni valores no presentes.
4. Revisar filtros, filas ocultas, valores vacíos y posibles duplicados.
5. Validar manualmente la afirmación con filtros u ordenación antes de presentarla como hallazgo.

## Limpieza

1. El instructor cierra el panel de Copilot o deja Excel en un estado neutro para la siguiente actividad.
2. El instructor confirma que no se agregaron fórmulas, columnas, comentarios ni conclusiones al libro fuente.
3. Si Excel solicita guardar cambios, el instructor selecciona **No guardar**, salvo que se hayan producido cambios administrativos ajenos a los datos y estén autorizados.
4. El archivo `C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx` debe conservarse disponible para la práctica posterior `02-00-01`.
5. Los participantes conservan únicamente sus notas sobre prompts, evidencia y validaciones observadas.

## Resumen

En esta demostración, los participantes observaron cómo el instructor utiliza Microsoft 365 Copilot en Excel para explorar una tabla de portafolio mediante prompts progresivos. Se mostraron consultas para resumir datos, analizar tendencias, identificar valores relevantes y solicitar apoyo para fórmulas, manteniendo la validación humana como requisito obligatorio.

Puntos clave para la práctica posterior:

- Un prompt efectivo especifica objetivo, datos relevantes, período, criterio y formato de respuesta.
- Copilot puede acelerar la exploración, pero sus respuestas deben verificarse contra la tabla.
- Un máximo, mínimo o valor atípico es una señal para investigar, no una decisión automática.
- Las fórmulas propuestas requieren revisión de referencias, errores y reglas de negocio.
- La falta de evidencia debe producir una respuesta de incertidumbre o una solicitud de validación, nunca una invención de datos.

**Recursos opcionales**

- Microsoft Support: [Obtener información sobre los datos con Copilot en Excel](https://support.microsoft.com/es-es/topic/obtener-informaci%C3%B3n-sobre-los-datos-con-copilot-en-excel-1bba8e8f-97f8-43c0-a89f-6c1e8c7f4f97)
- Microsoft Support: [Ayuda y aprendizaje de Copilot en Excel](https://support.microsoft.com/es-es/excel)
- Microsoft Learn: [Introducción a Microsoft 365 Copilot](https://learn.microsoft.com/es-es/copilot/microsoft-365/microsoft-365-copilot-overview)
