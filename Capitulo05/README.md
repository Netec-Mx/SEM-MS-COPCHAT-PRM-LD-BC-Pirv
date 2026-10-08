# Práctica guiada 4. Resumir un hilo de comité y preparar el correo ejecutivo de seguimiento con decisiones, pendientes y próximos pasos

## Metadatos

| Campo | Valor |
|---|---|
| Duración | **20 min** |
| Complejidad | Media |
| Nivel de Bloom | Aplicar + Analizar + Crear |
| Aplicación | Outlook con Microsoft 365 Copilot |
| Insumo principal | Hilo importado desde `recursos/correos_comite/` |
| Insumos de continuidad | Memo y deck de los capítulos anteriores |
| Resultado | Borrador de correo ejecutivo; **no se envía** |

## Descripción General

Utilizarás Copilot en Outlook para resumir una conversación de correo, comprobar el resumen contra los mensajes originales y redactar un seguimiento ejecutivo. El resultado permanecerá en **Borradores**.

> [!IMPORTANT]
> La función de resumen trabaja sobre conversaciones de correo. Por eso los cuatro mensajes `.eml` deben importarse al buzón **antes** de iniciar los 20 minutos de práctica. Utiliza [`Hilo_Comite_Inversion.docx`](../recursos/Hilo_Comite_Inversion.docx) como versión de consulta y sigue el [Setup Guide](../SETUP_GUIDE.md) para importar los correos.

## Objetivos de Aprendizaje

- [ ] Resumir un hilo con Copilot.
- [ ] Separar decisiones confirmadas, acciones, riesgos y asuntos pendientes.
- [ ] Comprobar la trazabilidad de cada afirmación contra el hilo.
- [ ] Generar un correo ejecutivo de seguimiento y conservarlo como borrador.

## Prerrequisitos

### Conocimientos requeridos

- Diferencia entre una propuesta, una decisión confirmada y un asunto pendiente.
- Revisión básica de correo en Outlook.

### Acceso requerido

- Outlook con Microsoft 365 Copilot.
- Los cuatro archivos `.eml` importados y visibles como una conversación en Outlook.
- Memo y deck disponibles como referencia de continuidad.

## Entorno de Laboratorio

Microsoft documenta que Copilot en Outlook puede resumir un hilo y generar un borrador de correo. La posición exacta de las opciones puede variar entre Outlook web, nuevo Outlook, clásico o Mac.

## Instrucciones Paso a Paso

### Paso 1: Delimitar la conversación fuente

**Tiempo:** 3 min  
**Objetivo:** asegurar que Copilot y el participante revisan el hilo correcto.

1. Abre Outlook.
2. Localiza el hilo **Comité de inversión | Revisión del portafolio y próximos pasos**.
3. Confirma que ves todos los mensajes de la conversación.
4. No abras otros hilos durante el análisis.

**Criterio de finalización:** identificaste una única conversación fuente.

---

### Paso 2: Generar el resumen con Copilot

**Tiempo:** 4 min  
**Objetivo:** obtener una síntesis inicial de la conversación.

1. Selecciona **Resumen por Copilot / Summarize** o la opción equivalente.
2. Lee el resultado completo.
3. Registra en tus notas cuatro grupos:
   - decisiones confirmadas;
   - acciones pendientes;
   - riesgos o dependencias;
   - asuntos por confirmar.

> [!NOTE]
> Si el resumen incluye citas o referencias al hilo, utilízalas para volver al mensaje fuente. Si tu interfaz no muestra citas, localiza manualmente el mensaje correspondiente.

**Criterio de finalización:** existe un resumen organizado en las cuatro categorías.

---

### Paso 3: Validar el resumen contra los mensajes

**Tiempo:** 4 min  
**Objetivo:** impedir que propuestas o ambigüedades se conviertan en decisiones.

Comprueba cada elemento del resumen:

| Pregunta de control | Acción |
|---|---|
| ¿Existe en un mensaje del hilo? | Conserva la referencia. |
| ¿Es una decisión explícita? | Márcala como confirmada. |
| ¿Es una propuesta o recomendación? | No la conviertas en decisión. |
| ¿Falta responsable o fecha? | Marca el dato como pendiente de confirmar. |
| ¿Hay contradicción? | Conserva la contradicción como tema por resolver. |

**Criterio de finalización:** todas las decisiones y acciones conservadas tienen evidencia en el hilo.

---

### Paso 4: Crear el correo ejecutivo de seguimiento

**Tiempo:** 6 min  
**Objetivo:** generar un borrador utilizable sin inventar compromisos.

1. Crea un correo nuevo.
2. Deja el campo **Para** vacío, salvo que exista un destinatario de prueba autorizado.
3. Abre **Borrador con Copilot / Draft with Copilot**.
4. Utiliza:

> **PROMPT 1 — CORREO DE SEGUIMIENTO**
>
> ```text
> Redacta un correo ejecutivo de seguimiento basado únicamente en esta lista validada:
>
> Decisiones confirmadas:
> [PEGAR]
>
> Acciones pendientes:
> [PEGAR]
>
> Riesgos o dependencias:
> [PEGAR]
>
> Asuntos por confirmar:
> [PEGAR]
>
> Estructura:
> - asunto sugerido;
> - contexto de una frase;
> - decisiones confirmadas;
> - acciones pendientes;
> - riesgos o dependencias;
> - asuntos por confirmar;
> - próximos pasos;
> - cierre breve.
>
> No inventes responsables, fechas, aprobaciones ni decisiones.
> Si un dato no está confirmado, mantenlo explícitamente como pendiente.
> Tono ejecutivo, directo y neutral.
> ```

5. Inserta o conserva el borrador.
6. Si necesitas continuidad con el caso, agrega manualmente **una sola frase** indicando que el memo y el deck están disponibles para la siguiente revisión; no inventes resultados nuevos.

**Criterio de finalización:** el correo separa hechos confirmados de pendientes.

---

### Paso 5: Guardar y comprobar el borrador

**Tiempo:** 3 min  
**Objetivo:** conservar el artefacto sin enviarlo.

1. Revisa nuevamente nombres, decisiones y acciones.
2. Guarda el mensaje.
3. Cierra el editor.
4. Abre **Borradores**.
5. Confirma que el mensaje está allí y **no** en Elementos enviados.

**Criterio de finalización:** existe un único borrador validado y no fue enviado.

## Validación y Pruebas

| # | Criterio | Estado |
|---:|---|:---:|
| 1 | El resumen se generó sobre el hilo correcto. | ☐ |
| 2 | Las decisiones confirmadas pueden localizarse en un mensaje fuente. | ☐ |
| 3 | Los responsables o fechas ausentes permanecen pendientes de confirmar. | ☐ |
| 4 | El correo separa decisiones, acciones, riesgos y pendientes. | ☐ |
| 5 | El mensaje permanece en Borradores. | ☐ |

## Solución de Problemas

| Situación | Qué hacer |
|---|---|
| No aparece Resumen por Copilot | Confirma licencia, cuenta y que abriste una conversación de correo compatible. |
| El resumen presenta una propuesta como decisión | Vuelve al mensaje fuente y corrige la lista validada antes de generar el correo. |
| No aparece Borrador con Copilot | Verifica que Copilot esté habilitado en Outlook; en nuevo Outlook/web, usa formato HTML si la función no está disponible en texto sin formato. |
| El hilo no aparece en el buzón | Importa previamente los cuatro `.eml` de `recursos/correos_comite/` siguiendo el Setup Guide. Si ya se importaron, revisa la carpeta de destino y la vista por conversación. |

## Limpieza

- Conserva el borrador para comparar el resultado después del Capítulo 6.
- No envíes el mensaje.
- No elimines el hilo de práctica hasta finalizar el curso.

## Resumen

Resumiste un hilo, verificaste cada afirmación contra los mensajes originales y generaste un correo ejecutivo de seguimiento. El borrador conserva explícitamente la diferencia entre decisiones confirmadas y asuntos pendientes.
