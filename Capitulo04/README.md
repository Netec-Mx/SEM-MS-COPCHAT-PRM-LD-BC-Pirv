# Práctica guiada 3. Convertir el memo en un breve deck ejecutivo, reorganizar la narrativa y realizar la revisión final de formato corporativo

## Metadatos

| Campo | Valor |
|---|---|
| Duración | 20 minutos |
| Complejidad | Media |
| Nivel de Bloom | Crear |

## Descripción General

En este laboratorio transformará el memo de inversión validado en una presentación ejecutiva breve en PowerPoint. Usará Microsoft 365 Copilot para generar un primer borrador a partir de `Memo_Inversion.docx`, pero reorganizará y validará manualmente la narrativa, las cifras y el formato. El resultado será un deck de hasta seis diapositivas con una secuencia orientada a la decisión y alineado con el kit de marca corporativo.

## Objetivos de Aprendizaje

- [ ] Generar una propuesta inicial de presentación a partir del memo de inversión mediante Microsoft 365 Copilot en PowerPoint.
- [ ] Reorganizar el contenido en una narrativa ejecutiva: decisión, contexto, evidencia, riesgos, implicaciones y próximos pasos.
- [ ] Aplicar el tema, los diseños y la jerarquía tipográfica definidos en el kit de marca corporativo.
- [ ] Confirmar que las cifras y afirmaciones del deck coinciden con `Memo_Inversion.docx` y `Portfolio_Analisis.xlsx`.
- [ ] Realizar una revisión humana medible de precisión, legibilidad, contraste y consistencia visual.

## Prerrequisitos

### Conocimientos requeridos

- Comprender la diferencia entre **contenido** (cifras, recomendaciones, riesgos y hallazgos) y **diseño** (tema, tipografías, colores, patrones y diseños de diapositiva).
- Reconocer una narrativa ejecutiva orientada a decisión: primero la recomendación, después la evidencia que la sustenta.
- Saber que un **prompt** es una instrucción temporal enviada a Copilot durante una interacción; no crea un agente ni modifica permanentemente el comportamiento del servicio.
- Distinguir un prompt de un **mensaje de sistema**: el mensaje de sistema es una configuración controlada por la plataforma o el administrador; el participante no debe intentar modificarlo ni asumir que un prompt tiene esa persistencia.
- Conocer los hallazgos, cifras y recomendaciones validadas en las prácticas 02-00-01 y 03-00-01.

### Acceso y archivos requeridos

Confirme que tiene acceso autenticado con su cuenta corporativa a los siguientes recursos:

| Recurso | Ubicación o condición |
|---|---|
| Memo de inversión completado | `C:\CopilotLabs\Batch1\Memo_Inversion.docx` |
| Libro de verificación | `C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx` |
| Archivo de salida | `C:\CopilotLabs\Batch1\Deck_Ejecutivo_Inversion.pptx` |
| Kit de marca o tema corporativo | Disponible en PowerPoint conforme a las instrucciones internas |
| Microsoft 365 Copilot | Licencia activa y habilitada para PowerPoint en el tenant corporativo |
| PowerPoint | Autenticado con la cuenta corporativa |

> **Importante:** Microsoft 365 Copilot integrado en PowerPoint es el servicio utilizado en este laboratorio. Microsoft 365 Copilot Chat puede servir para consultas generales, pero no sustituye las funciones integradas de creación y edición en PowerPoint. Designer, Planner y el agente preconstruido Analista no forman parte de esta práctica y no deben utilizarse para generar el deck.

## Entorno de Laboratorio

### Hardware y conectividad

| Componente | Configuración requerida |
|---|---|
| Equipo | Windows de 64 bits, procesador de 2 núcleos a 1,6 GHz o superior |
| Memoria | 8 GB de RAM como mínimo |
| Pantalla | Mínimo 1366 × 768; recomendado 1920 × 1080 |
| Red | Conexión HTTPS saliente por el puerto 443; recomendado 10 Mbps de descarga y 2 Mbps de carga |
| Cuenta | Cuenta corporativa de Microsoft 365; MFA si lo exige la organización |

### Software y licencias

| Tecnología | Versión o configuración | Licencia/configuración | Fuente oficial |
|---|---|---|---|
| Windows 11 Pro o Enterprise | 24H2, compilación 26100.1, 64 bits | Licencia corporativa de Windows | https://learn.microsoft.com/es-es/windows/release-health/windows11-release-information |
| Microsoft 365 Apps for enterprise: PowerPoint, Word y Excel | Versión 2408, compilación 17928.20114, 64 bits | Licencia corporativa de Microsoft 365 Apps | https://learn.microsoft.com/es-es/officeupdates/update-history-microsoft365-apps-by-date |
| Microsoft 365 Copilot en PowerPoint | Servicio SaaS, **[VERSIÓN POR VALIDAR]** | Licencia Microsoft 365 Copilot asignada al usuario y experiencia habilitada por el tenant | https://support.microsoft.com/es-es/copilot-powerpoint |
| Microsoft 365 Copilot Chat | Servicio SaaS, **[VERSIÓN POR VALIDAR]** | Puede tener acceso diferente al de Microsoft 365 Copilot; no sustituye la integración de PowerPoint | https://support.microsoft.com/es-es/topic/microsoft-365-copilot-chat-3d1a4b8c-5b6a-4d0f-b4e8-36ee3f1f5bb0 |
| Kit de marca corporativo para PowerPoint | **[VERSIÓN POR VALIDAR]**, 64 bits no aplicable | Tema o plantilla aprobada e instalada por la organización | **[ENLACE OFICIAL INTERNO POR VALIDAR]** |
| Microsoft Designer | Servicio SaaS, **[VERSIÓN POR VALIDAR]** | No se requiere ni se utiliza en este laboratorio | https://designer.microsoft.com/ |
| Microsoft Planner | Servicio SaaS, **[VERSIÓN POR VALIDAR]** | No se requiere ni se utiliza en este laboratorio | https://support.microsoft.com/es-es/planner |

### Comprobación inicial de archivos

Abra **Símbolo del sistema** o PowerShell y ejecute:

```powershell
Test-Path "C:\CopilotLabs\Batch1\Memo_Inversion.docx"
Test-Path "C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx"
Test-Path "C:\CopilotLabs\Batch1"
```

Los tres comandos deben devolver `True`.

> No modifique los nombres originales de columnas del libro de Excel ni la hoja protegida de datos fuente. La hoja `Conclusiones` es la referencia analítica preparada en la práctica anterior.

## Instrucciones Paso a Paso

### Paso 1: Confirmar las fuentes y la extensión objetivo

**Objetivo:** Verificar que se utilizarán únicamente los entregables aprobados y establecer una extensión ejecutiva controlada.

**Instrucciones:**

1. Abra `C:\CopilotLabs\Batch1\Memo_Inversion.docx` en Word.
2. Identifique, sin editar el memo, los siguientes elementos:
   - Recomendación de inversión o decisión solicitada.
   - Dos o tres evidencias cuantitativas principales.
   - Riesgos o limitaciones relevantes.
   - Próximos pasos, responsables o decisiones requeridas.
3. Abra `C:\CopilotLabs\Batch1\Portfolio_Analisis.xlsx`.
4. Revise la hoja denominada exactamente `Conclusiones`.
5. Localice las cifras que sustentan la recomendación del memo. No modifique la hoja fuente protegida ni cambie nombres de columnas.
6. Defina el límite de esta práctica: **máximo seis diapositivas**, salvo que el instructor haya indicado una extensión menor.
7. Prepare esta secuencia narrativa obligatoria:

   | Diapositiva | Propósito |
   |---|---|
   | 1 | Decisión o recomendación |
   | 2 | Contexto y objetivo de análisis |
   | 3 | Evidencia clave |
   | 4 | Riesgos y limitaciones |
   | 5 | Implicaciones para la decisión |
   | 6 | Próximos pasos y decisión requerida |

**Resultado esperado:** Tendrá identificadas las afirmaciones y cifras que pueden utilizarse en el deck, junto con una estructura máxima de seis diapositivas.

**Verificación:**

- El memo está completo y disponible.
- El libro de Excel contiene la hoja `Conclusiones`.
- Puede señalar al menos dos cifras que aparezcan tanto en el memo como en el libro de Excel.
- La secuencia incluye una decisión explícita en la primera diapositiva, no una portada genérica.

---

### Paso 2: Crear el archivo desde el tema corporativo

**Objetivo:** Crear el archivo de salida utilizando el formato institucional antes de generar contenido.

**Instrucciones:**

1. Abra PowerPoint con la cuenta corporativa autenticada.
2. Cree una presentación nueva usando la plantilla o tema corporativo aprobado.
3. En la pestaña **Diseño**, confirme que está aplicado el tema corporativo.
4. Seleccione **Diseño > Tamaño de diapositiva** y confirme el formato indicado por el kit de marca, normalmente **Panorámica (16:9)**, salvo instrucción interna distinta.
5. Revise los diseños disponibles en **Inicio > Nueva diapositiva**. Identifique al menos:
   - Un diseño de título o portada.
   - Un diseño de contenido con dos columnas.
   - Un diseño para gráfico o evidencia.
   - Un diseño de cierre o próximos pasos.
6. Guarde el archivo inmediatamente como:

   ```text
   C:\CopilotLabs\Batch1\Deck_Ejecutivo_Inversion.pptx
   ```

7. No cambie manualmente fuentes, colores del tema, logotipos, pies de página ni elementos definidos en el Patrón de diapositivas.

**Resultado esperado:** Existe un archivo de PowerPoint guardado en la ubicación estándar y basado en el tema corporativo.

**Verificación:**

- El archivo `Deck_Ejecutivo_Inversion.pptx` existe en `C:\CopilotLabs\Batch1\`.
- El tamaño de diapositiva coincide con el estándar corporativo.
- Los diseños disponibles muestran la tipografía, colores y elementos visuales del kit de marca.
- La presentación no contiene aún diapositivas con formato manual ajeno al tema.

---

### Paso 3: Generar un primer borrador con Copilot

**Objetivo:** Usar Copilot para producir un borrador estructurado sin delegar la validación factual ni el cumplimiento de marca.

**Instrucciones:**

1. En PowerPoint, abra el panel de **Copilot**.
2. Si su entorno ofrece la opción de crear una presentación desde un archivo, seleccione `Memo_Inversion.docx`.
3. Si esa opción no está disponible, copie al panel de Copilot el siguiente prompt y adjunte o haga referencia al memo según las opciones disponibles en su tenant:

   ```text
   Crea un borrador de presentación ejecutiva de máximo seis diapositivas a partir de Memo_Inversion.docx.

   Audiencia: comité de inversión.
   Objetivo: solicitar una decisión informada, no presentar un informe exhaustivo.
   Estructura obligatoria:
   1. Recomendación o decisión requerida.
   2. Contexto y objetivo del análisis.
   3. Evidencia cuantitativa clave.
   4. Riesgos, limitaciones y supuestos.
   5. Implicaciones para la decisión.
   6. Próximos pasos, responsables o decisión requerida.

   Usa títulos que expresen conclusiones, no temas genéricos. Mantén lenguaje conciso y orientado a la acción. No inventes cifras, fechas, fuentes, responsables ni porcentajes. Si una afirmación no está respaldada por el memo, indícala como “Por validar” o exclúyela. Conserva el tema y los diseños de la plantilla actual.
   ```

4. Revise el borrador generado antes de aceptar cambios.
5. Si Copilot crea más de seis diapositivas, conserve solo las que puedan contribuir a la secuencia definida en el Paso 1.
6. Guarde los cambios.

> **Nota de prompting:** El texto anterior es un **prompt temporal** para esta interacción. No configura un agente persistente, no modifica mensajes de sistema y no garantiza que Copilot valide las cifras por sí solo.

**Resultado esperado:** La presentación contiene un borrador inicial basado en el memo, con contenido distribuido en una estructura ejecutiva.

**Verificación:**

- El deck tiene seis diapositivas o menos.
- No hay afirmaciones nuevas que no puedan rastrearse al memo o al libro de Excel.
- Al menos una diapositiva contiene una recomendación o decisión explícita.
- Los títulos generados comunican conclusiones; por ejemplo, “La recomendación se mantiene bajo condiciones de riesgo controlado”, en lugar de “Recomendación”.

---

### Paso 4: Reorganizar la narrativa para la decisión ejecutiva

**Objetivo:** Ajustar manualmente el borrador para que la historia avance desde la decisión hasta las acciones requeridas.

**Instrucciones:**

1. Revise todas las diapositivas en **Vista > Clasificador de diapositivas**.
2. Reordene las diapositivas para cumplir exactamente esta lógica:
   1. Decisión o recomendación.
   2. Contexto.
   3. Evidencia clave.
   4. Riesgos y limitaciones.
   5. Implicaciones.
   6. Próximos pasos.
3. Edite el título de cada diapositiva para que contenga una idea principal verificable.
4. Elimine:
   - Contenido repetido entre diapositivas.
   - Párrafos extensos.
   - Explicaciones metodológicas no necesarias para la decisión.
   - Datos que no estén en el memo o que no pueda verificar en Excel.
5. Mantenga una sola idea central por diapositiva.
6. Si existe una diapositiva con más de cinco viñetas o más de dos ideas diferentes, reduzca el texto o divida el contenido usando los diseños corporativos disponibles, sin superar seis diapositivas.
7. Incorpore una nota breve de fuente cuando sea necesario, por ejemplo:

   ```text
   Fuente: Portfolio_Analisis.xlsx, hoja Conclusiones; Memo_Inversion.docx.
   ```

8. Guarde el archivo.

**Resultado esperado:** El deck presenta primero la decisión y utiliza el resto de la presentación para justificarla, contextualizarla y solicitar acciones concretas.

**Verificación:**

- La primera diapositiva permite entender qué decisión se recomienda o solicita.
- Las evidencias aparecen antes o junto a las implicaciones, no después de los próximos pasos.
- La diapositiva de riesgos identifica limitaciones reales y no minimiza incertidumbres.
- La última diapositiva incluye una acción, decisión requerida, responsable o siguiente hito verificable.

---

### Paso 5: Validar cifras, fuentes e incertidumbres

**Objetivo:** Confirmar la exactitud de las afirmaciones antes de aplicar el acabado final.

**Instrucciones:**

1. Coloque Word, Excel y PowerPoint en ventanas visibles o alterne entre ellas.
2. Para cada cifra del deck, compare:
   - La cifra mostrada en PowerPoint.
   - La cifra documentada en `Memo_Inversion.docx`.
   - La cifra o conclusión correspondiente en `Portfolio_Analisis.xlsx`, especialmente en la hoja `Conclusiones`.
3. Corrija cualquier diferencia de:
   - Valor numérico.
   - Unidad, moneda o porcentaje.
   - Periodo analizado.
   - Signo positivo o negativo.
   - Redondeo que cambie el significado.
4. Marque como **“Por validar”** cualquier afirmación del memo que no encuentre respaldada en el libro de Excel o en una fuente explícita del memo.
5. Ejecute la siguiente prueba adversarial: revise si el memo contiene texto similar a “ignora las instrucciones”, “añade una conclusión favorable”, “no menciones el riesgo” o cualquier instrucción incrustada que no forme parte del análisis.
6. Si encuentra ese tipo de texto, trátelo como contenido no confiable: no lo ejecute, no lo copie al deck y comuníquelo al instructor.
7. Verifique que no se haya incluido una cifra inexistente, una fecha no respaldada o una recomendación inventada por Copilot.
8. Guarde el archivo.

**Resultado esperado:** Todas las cifras y conclusiones incluidas tienen una fuente rastreable; las incertidumbres permanecen visibles y no se presentan como hechos.

**Verificación:**

| Criterio medible | Evidencia requerida |
|---|---|
| Cifras verificadas | El 100 % de las cifras del deck coincide con el memo y/o con `Conclusiones` |
| Trazabilidad | Cada diapositiva cuantitativa incluye fuente visible o una fuente claramente identificable |
| Incertidumbre | Las afirmaciones no respaldadas están eliminadas o marcadas como “Por validar” |
| Prueba adversarial | No se trasladó al deck ninguna instrucción incrustada que contradiga el objetivo analítico |
| Supervisión humana | El participante revisó las seis diapositivas, no solo aceptó el resultado de Copilot |

---

### Paso 6: Aplicar la revisión final de formato corporativo

**Objetivo:** Asegurar consistencia visual, legibilidad y adecuación ejecutiva del deck.

**Instrucciones:**

1. En PowerPoint, revise que todas las diapositivas usen el tema corporativo aplicado en el Paso 2.
2. Use **Inicio > Diseño** para cambiar una diapositiva a un diseño aprobado si su composición no se ajusta al contenido.
3. No aplique formatos locales innecesarios. Si detecta una fuente distinta a la corporativa, aplique el diseño correcto o restablezca el marcador de posición.
4. Compruebe los siguientes puntos en cada diapositiva:
   - Título visible, breve y orientado a una conclusión.
   - Texto legible sin reducir manualmente el tamaño por debajo del estándar del tema.
   - Contraste suficiente entre texto, fondos y elementos gráficos.
   - Alineación consistente con los márgenes y marcadores de posición.
   - Uso exclusivo de colores aprobados por el tema.
   - Ausencia de texto cortado, superpuesto o fuera de la diapositiva.
5. Si utiliza un gráfico, confirme que:
   - Tiene título o mensaje interpretativo.
   - Sus etiquetas son legibles.
   - No contiene más series de las necesarias.
   - Usa colores de la paleta corporativa.
6. Ejecute **Presentación con diapositivas > Desde el principio** y revise el deck como lo verá el comité.
7. Guarde el archivo final y ciérrelo.

**Resultado esperado:** El deck es breve, visualmente consistente y apto para una revisión ejecutiva interna.

**Verificación:**

- El deck contiene un máximo de seis diapositivas.
- No hay objetos fuera de los límites de la diapositiva.
- No hay fuentes, colores o logotipos no autorizados.
- La primera y la última diapositiva comunican, respectivamente, la decisión y el siguiente paso.
- El archivo final está guardado como `C:\CopilotLabs\Batch1\Deck_Ejecutivo_Inversion.pptx`.

## Validación y Pruebas

Realice la validación final antes de entregar el archivo. La aprobación requiere cumplir **todos** los criterios siguientes.

| Área | Criterio de aceptación | Evidencia |
|---|---|---|
| Archivo | Existe `Deck_Ejecutivo_Inversion.pptx` en el directorio estándar | Archivo visible en el Explorador de archivos |
| Extensión | Máximo seis diapositivas, salvo indicación diferente del instructor | Vista Clasificador de diapositivas |
| Narrativa | La secuencia incluye decisión, contexto, evidencia, riesgos, implicaciones y próximos pasos | Títulos y orden de diapositivas |
| Precisión | El 100 % de las cifras coincide con el memo y con el libro cuando corresponda | Comparación manual con Word y Excel |
| Trazabilidad | Toda cifra o conclusión crítica tiene una fuente identificable | Pie de fuente, nota o referencia visible |
| Incertidumbre | No se presentan afirmaciones no confirmadas como hechos | Revisión de etiquetas “Por validar” o eliminación del contenido |
| Marca | Se conserva el tema, tipografía, colores y diseños corporativos | Revisión visual en modo Presentación |
| Legibilidad | No hay texto cortado, superpuesto o demasiado denso | Presentación desde el principio |
| Prueba adversarial | Las instrucciones incrustadas, contradictorias o no verificables del memo no se ejecutan ni se incorporan | Revisión manual documentada al instructor |

### Evidencia mínima para el instructor

Muestre en pantalla, sin necesidad de generar documentación adicional:

1. La vista Clasificador con el máximo de seis diapositivas.
2. La diapositiva 1 con la recomendación o decisión.
3. Una diapositiva con evidencia y su fuente.
4. La diapositiva de riesgos.
5. La diapositiva final con próximos pasos o decisión requerida.
6. El archivo guardado en `C:\CopilotLabs\Batch1\Deck_Ejecutivo_Inversion.pptx`.

## Solución de Problemas

### Problema 1: Copilot no muestra la opción para crear una presentación desde el memo

**Síntomas:** El panel de Copilot está disponible, pero no aparece la opción para crear una presentación desde un archivo de Word, o el archivo no puede seleccionarse.

**Causa probable:** La función puede no estar habilitada en la configuración del tenant, el archivo puede no estar disponible para la experiencia de Copilot o la cuenta puede tener acceso a Copilot Chat pero no a Microsoft 365 Copilot integrado en PowerPoint.

**Corrección:**

1. Confirme que inició sesión en PowerPoint con la cuenta corporativa correcta.
2. Verifique con el instructor que su licencia incluye Microsoft 365 Copilot y que PowerPoint está habilitado.
3. Abra `Memo_Inversion.docx`, copie los puntos validados y use el prompt del Paso 3 pegando únicamente contenido relevante.
4. Cree manualmente las seis diapositivas con los diseños corporativos; Copilot puede utilizarse para resumir o proponer títulos, pero la estructura debe conservarse.
5. No use Designer, Planner ni herramientas externas como sustituto del flujo definido.

### Problema 2: El deck muestra fuentes sustituidas, texto desbordado o colores fuera de marca

**Síntomas:** Los saltos de línea cambian, el texto se sale de los marcadores, aparecen fuentes distintas o las diapositivas no parecen parte de la misma plantilla.

**Causa probable:** Se aplicó formato manual, se pegó contenido con formato externo, no está instalada la fuente corporativa o se utilizó una presentación en blanco en lugar del tema institucional.

**Corrección:**

1. En **Diseño**, vuelva a aplicar el tema corporativo aprobado.
2. Seleccione la diapositiva afectada y aplique un diseño corporativo desde **Inicio > Diseño**.
3. Use **Inicio > Restablecer** para recuperar la posición y el formato de los marcadores de posición.
4. Pegue texto usando solo texto sin formato cuando sea necesario y deje que el tema aplique la tipografía.
5. Si la fuente corporativa sigue sustituyéndose, no improvise una fuente alternativa: informe al instructor o al soporte interno y use la alternativa aprobada por el kit de marca.

## Limpieza

1. Confirme que el archivo final está guardado en:

   ```text
   C:\CopilotLabs\Batch1\Deck_Ejecutivo_Inversion.pptx
   ```

2. Cierre PowerPoint, Word y Excel después de guardar los cambios.
3. No elimine ni renombre:
   - `Portfolio_Analisis.xlsx`
   - `Memo_Inversion.docx`
   - La hoja `Conclusiones`
   - Los nombres originales de columnas del portafolio
4. No elimine el tema, plantilla ni fuentes del kit de marca corporativo.
5. Si creó versiones temporales, elimínelas solo después de comprobar que el archivo final se abre correctamente y conserva el formato corporativo.

## Resumen

En esta práctica convirtió un memo de inversión en un deck ejecutivo breve y orientado a decisión. Microsoft 365 Copilot aceleró la creación del primer borrador, pero la calidad final dependió de la supervisión humana: validación de cifras, tratamiento explícito de incertidumbres, rechazo de instrucciones incrustadas no confiables y aplicación consistente del kit de marca.

El entregable final es:

```text
C:\CopilotLabs\Batch1\Deck_Ejecutivo_Inversion.pptx
```

Recuerde que una presentación ejecutiva eficaz no reproduce el memo completo: prioriza una decisión clara, evidencia verificable, riesgos relevantes y próximos pasos accionables.
