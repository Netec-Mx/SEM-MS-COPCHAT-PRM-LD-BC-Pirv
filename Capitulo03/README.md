# Práctica guiada 2. Crear y refinar un memo de inversión utilizando los hallazgos del análisis realizado en Excel

## Metadatos

| Campo | Valor |
|---|---|
| Duración | **20 min** |
| Complejidad | Media |
| Nivel de Bloom | Aplicar + Crear |
| Aplicación | Word con Microsoft 365 Copilot |
| Insumo | Hoja `Conclusiones` del libro del Capítulo 2 |
| Resultado | `Memo_Inversion_[Iniciales].docx` |

## Descripción General

Transformarás las conclusiones **ya verificadas en Excel** en un memo ejecutivo. Copilot se utilizará para generar un primer borrador, reescribir, resumir y revisar la coherencia, pero cada cifra y afirmación relevante deberá mantenerse trazable al libro.

## Objetivos de Aprendizaje

- [ ] Crear un memo desde hallazgos verificados.
- [ ] Usar prompts para redactar, resumir y reescribir contenido en Word.
- [ ] Refinar tono, longitud y jerarquía para una audiencia ejecutiva.
- [ ] Mantener trazabilidad entre el memo y Excel.

## Prerrequisitos

### Conocimientos requeridos

- Haber completado el Capítulo 2.
- Reconocer qué conclusiones están verificadas y cuáles permanecen pendientes.

### Acceso requerido

- Word con Microsoft 365 Copilot.
- `Portfolio_Analisis_[Iniciales].xlsx` con la hoja `Conclusiones` completa.
- Carpeta de trabajo en OneDrive o SharePoint.

## Entorno de Laboratorio

No necesitas una plantilla `.docx` precreada. La práctica parte de un documento en blanco y genera el artefacto dentro del tiempo asignado.

## Instrucciones Paso a Paso

### Paso 1: Preparar el contenido fuente

**Tiempo:** 3 min  
**Objetivo:** transferir a Word solo información validada.

1. Abre `Portfolio_Analisis_[Iniciales].xlsx` y la hoja **Conclusiones**.
2. Selecciona entre 3 y 5 conclusiones verificadas.
3. Copia a Word únicamente estas columnas o su equivalente:
   - hallazgo;
   - evidencia;
   - aspecto pendiente.
4. Crea un documento en blanco.
5. Guárdalo en tu carpeta de trabajo como:

```text
Memo_Inversion_[Iniciales].docx
```

**Criterio de finalización:** Word contiene el material fuente validado y el archivo ya está guardado.

---

### Paso 2: Generar el primer borrador

**Tiempo:** 5 min  
**Objetivo:** convertir los hallazgos en una estructura ejecutiva.

Abre Copilot en Word y usa:

> **PROMPT 1 — BORRADOR DEL MEMO**
>
> ```text
> A partir únicamente del contenido validado que aparece en este documento, crea un memo ejecutivo de inversión.
>
> Estructura:
> 1. Propósito.
> 2. Contexto.
> 3. Hallazgos principales.
> 4. Riesgos o puntos de atención.
> 5. Recomendación preliminar, solo si está respaldada por el contenido fuente.
> 6. Próximos pasos.
> 7. Aspectos pendientes de confirmar.
>
> Reglas:
> - No inventes cifras, causas, responsables ni decisiones.
> - Si falta información, indícala como pendiente de confirmar.
> - Mantén tono ejecutivo y lenguaje claro.
> - No elimines la trazabilidad hacia la evidencia original.
> ```

Revisa el borrador antes de conservarlo.

**Criterio de finalización:** existe un memo estructurado sin datos nuevos no sustentados.

---

### Paso 3: Refinar para audiencia ejecutiva

**Tiempo:** 4 min  
**Objetivo:** mejorar claridad y jerarquía.

Selecciona el cuerpo del memo y solicita:

> **PROMPT 2 — REFINAMIENTO EJECUTIVO**
>
> ```text
> Reescribe este memo para una audiencia ejecutiva.
> Mantén todos los hechos y cifras sin alterarlos.
> Prioriza primero los hallazgos de mayor impacto, luego los riesgos y finalmente los próximos pasos.
> Reduce repeticiones y evita lenguaje promocional.
> ```

Compara la nueva versión con el contenido fuente.

**Criterio de finalización:** el memo es más conciso sin perder evidencia ni introducir nuevos hechos.

---

### Paso 4: Resumir y consultar coherencia

**Tiempo:** 4 min  
**Objetivo:** comprobar que el mensaje principal coincide con el documento.

Envía:

> **PROMPT 3 — CONTROL DE COHERENCIA**
>
> ```text
> Resume este memo en cinco viñetas.
> Después identifica:
> - cualquier cifra sin evidencia visible en el documento;
> - afirmaciones que parezcan más fuertes que la evidencia;
> - información que debería permanecer como pendiente de confirmar.
> No corrijas nada todavía; solo señala los puntos a revisar.
> ```

Corrige manualmente lo necesario.

**Criterio de finalización:** el resumen representa fielmente el memo y no quedan afirmaciones no justificadas.

---

### Paso 5: Preparar el documento para PowerPoint

**Tiempo:** 4 min  
**Objetivo:** dejar una fuente clara para el siguiente capítulo.

1. Aplica estilos de Word a los encabezados principales.
2. Verifica que cada sección tenga un título claro.
3. Confirma que el documento se guardó y sincronizó.
4. Mantén el archivo en OneDrive o SharePoint para facilitar su referencia desde PowerPoint.

> [!TIP]
> Microsoft recomienda usar estilos de Word cuando se va a crear una presentación desde un documento, porque ayudan a Copilot a interpretar la estructura.

**Criterio de finalización:** `Memo_Inversion_[Iniciales].docx` está guardado, estructurado y listo para usarse como fuente.

## Validación y Pruebas

| # | Criterio | Estado |
|---:|---|:---:|
| 1 | El memo contiene propósito, hallazgos, riesgos y próximos pasos. | ☐ |
| 2 | Las cifras conservadas pueden rastrearse a `Conclusiones`. | ☐ |
| 3 | Los asuntos no confirmados están marcados como tales. | ☐ |
| 4 | Se realizó al menos una reescritura con Copilot. | ☐ |
| 5 | El archivo está guardado en OneDrive o SharePoint. | ☐ |

## Solución de Problemas

| Situación | Qué hacer |
|---|---|
| Copilot agrega hechos nuevos | Elimínalos y vuelve a pedir que use exclusivamente el contenido del documento. |
| El memo es demasiado largo | Selecciona la sección y solicita una versión más concisa sin cambiar cifras ni significado. |
| Copilot no encuentra el archivo | Confirma que el documento está guardado y sincronizado con OneDrive o SharePoint. |

## Limpieza

- Conserva el memo: será el insumo del Capítulo 4.
- Cierra Excel si ya no lo necesitas, pero conserva el libro para el Capítulo 6.

## Resumen

Convertiste hallazgos verificados en un memo ejecutivo, lo refinaste con Copilot y comprobaste su coherencia antes de utilizarlo como fuente de PowerPoint.