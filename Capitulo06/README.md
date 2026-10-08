# Práctica guiada 5. Consultar el agente Analista sobre el mismo caso, refinar la solicitud y reutilizar su resultado para complementar el memo o el correo final

## Metadatos

| Campo | Valor |
|---|---|
| Duración de práctica | **25 min** |
| Bloque 6.1 previo | **5 min** |
| Tiempo total del capítulo | **30 min** |
| Complejidad | Media |
| Nivel de Bloom | Analizar + Evaluar + Crear |
| Aplicación | Microsoft Copilot — agente Analista / Analyst |
| Insumo principal | `Portfolio_Analisis_[Iniciales].xlsx` |
| Resultado | Hallazgos adicionales validados e incorporados al memo o al borrador |

## Descripción General

Usarás el agente **Analista / Analyst** sobre el mismo libro del caso para ampliar el análisis. Después compararás sus resultados con Excel y reutilizarás solo los hallazgos comprobados en el memo o en el correo del capítulo anterior.

> [!IMPORTANT]
> Analista es un agente creado por Microsoft orientado al análisis de datos. Microsoft indica que puede trabajar con archivos como Excel o CSV, calcular estadísticas, identificar tendencias y valores atípicos y devolver informes con tablas o gráficos. Si no está disponible, puede deberse a configuración o administración del tenant.

## Objetivos de Aprendizaje

- [ ] Diferenciar un agente preconstruido de un agente declarativo.
- [ ] Adjuntar el libro del caso a Analista y formular una solicitud acotada.
- [ ] Refinar el prompt para exigir evidencia y distinguir hechos de interpretaciones.
- [ ] Validar los hallazgos antes de incorporarlos a otro artefacto.

## Prerrequisitos

### Conocimientos requeridos

- Haber completado los capítulos anteriores.
- Comprender que un agente especializado no elimina la necesidad de validar la fuente.

### Acceso requerido

- Microsoft Copilot con **Analista / Analyst** habilitado.
- `Portfolio_Analisis_[Iniciales].xlsx` guardado y accesible desde OneDrive, SharePoint o el dispositivo, según las opciones de adjuntar contenido disponibles.
- Memo y/o borrador de Outlook disponibles para actualización.

## Entorno de Laboratorio

### Bloque 6.1 — Ubicar agentes y reconocer su propósito — 5 min

Antes de la práctica:

1. Abre Microsoft Copilot.
2. Localiza **Agentes / Agents**.
3. Abre **Analista / Analyst** y revisa su descripción.
4. Distingue:
   - **Agente preconstruido:** experiencia creada y publicada por Microsoft para una tarea especializada, como Analista.
   - **Agente declarativo:** agente configurado para un escenario específico mediante propósito, instrucciones, conocimiento y acciones.
5. No crees ni configures un agente durante este curso.

## Instrucciones Paso a Paso

### Paso 1: Adjuntar la fuente correcta

**Tiempo:** 4 min  
**Objetivo:** asegurar que el agente trabaja sobre el mismo caso.

1. Abre **Analista / Analyst**.
2. Usa **Adjuntar contenido / Attach content**, `+` o la opción equivalente.
3. Selecciona `Portfolio_Analisis_[Iniciales].xlsx`.
4. Confirma que el archivo aparece asociado a la conversación.

> [!WARNING]
> No adjuntes documentos que no formen parte del laboratorio.

**Criterio de finalización:** el agente tiene acceso al libro correcto.

---

### Paso 2: Formular la primera consulta analítica

**Tiempo:** 7 min  
**Objetivo:** obtener un análisis especializado pero verificable.

Envía:

> **PROMPT 1 — ANÁLISIS DEL PORTAFOLIO**
>
> ```text
> Analiza el archivo adjunto del caso.
>
> Identifica hasta cinco hallazgos relevantes sobre:
> - tendencias entre períodos;
> - concentración;
> - rendimiento;
> - valores atípicos;
> - nivel de riesgo.
>
> Para cada hallazgo entrega:
> 1. Hecho observado.
> 2. Datos o columnas utilizados.
> 3. Evidencia numérica que pueda comprobar en Excel.
> 4. Interpretación posible.
> 5. Limitación o información faltante.
>
> No inventes causas ni decisiones de inversión.
> ```

Espera el resultado y localiza al menos dos afirmaciones verificables.

**Criterio de finalización:** la respuesta incluye evidencia o referencias comprobables en el archivo.

---

### Paso 3: Refinar la solicitud

**Tiempo:** 5 min  
**Objetivo:** reducir ambigüedad y priorizar lo que puede respaldarse.

Envía:

> **PROMPT 2 — REFINAMIENTO**
>
> ```text
> Revisa tu análisis anterior y conserva únicamente los tres hallazgos con evidencia más clara dentro del archivo.
>
> Ordénalos por impacto potencial, pero separa explícitamente:
> - evidencia;
> - interpretación;
> - dato pendiente de confirmar.
>
> Si no puedes sustentar un punto con el archivo, elimínalo del listado final.
> ```

Compara el nuevo resultado con el anterior.

**Criterio de finalización:** quedan como máximo tres hallazgos con trazabilidad clara.

---

### Paso 4: Validar y reutilizar el resultado

**Tiempo:** 6 min  
**Objetivo:** incorporar solo contenido comprobado al flujo de Microsoft 365.

1. Abre el libro en Excel.
2. Verifica al menos **dos** de los hallazgos del agente.
3. Selecciona solo los hallazgos que coincidan con la fuente.
4. Elige uno de estos destinos:
   - `Memo_Inversion_[Iniciales].docx`; o
   - el correo guardado en **Borradores**.
5. Agrega una sección breve titulada:

```text
Análisis complementario validado
```

6. Para cada punto incorporado, indica que fue **validado contra el libro**.
7. Si el resultado contradice una conclusión anterior, no ocultes la contradicción; registra que debe revisarse.

**Criterio de finalización:** el artefacto final incluye solo hallazgos que también pudieron comprobarse en Excel.

---

### Paso 5: Comprobación final

**Tiempo:** 3 min  
**Objetivo:** cerrar el flujo sin presentar la IA como fuente de verdad.

Revisa:

- que el archivo adjunto era el correcto;
- que los hallazgos reutilizados tienen evidencia;
- que una interpretación no se convirtió en hecho;
- que el correo, si fue el destino, sigue sin enviarse;
- que el memo o borrador se guardó correctamente.

**Criterio de finalización:** puedes explicar qué aportó Analista y cómo verificaste cada punto conservado.

## Validación y Pruebas

| # | Criterio | Estado |
|---:|---|:---:|
| 1 | La conversación se realizó con Analista / Analyst. | ☐ |
| 2 | Se adjuntó el libro correcto. | ☐ |
| 3 | El análisis se refinó a un máximo de tres hallazgos. | ☐ |
| 4 | Se comprobaron al menos dos hallazgos en Excel. | ☐ |
| 5 | Solo se reutilizó contenido validado. | ☐ |
| 6 | No se presentó una recomendación del agente como aprobación humana. | ☐ |

## Solución de Problemas

| Situación | Qué hacer |
|---|---|
| Analista no aparece | Confirma con el administrador que el agente está habilitado para tu cuenta o tenant. |
| No puedes adjuntar el libro | Usa la opción de carga desde dispositivo o selecciona la copia almacenada en OneDrive, según lo que ofrezca tu interfaz. |
| El agente devuelve una afirmación sin evidencia | Pide que indique columnas/datos concretos o descarta el punto. |
| El análisis contradice Excel | Excel y la fuente validada prevalecen; registra la contradicción y no reutilices el hallazgo. |

## Limpieza

- Conserva el libro, memo, deck y borrador si forman parte de la evidencia del curso.
- No envíes el correo de práctica.
- Cierra la conversación si trabajas en un equipo compartido.

## Resumen

Utilizaste el agente Analista para ampliar el análisis del mismo caso, refinaste la solicitud para exigir evidencia y reutilizaste únicamente hallazgos comprobados en Excel. El valor del agente se integra al flujo sin sustituir la validación humana.
