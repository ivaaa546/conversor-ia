import os
import zipfile

def create_docx(filename, title, subtitle, sections):
    # Minimal OpenXML structures
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

    doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:color w:val="2D3748"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:line="276" w:lineRule="auto" w:after="160"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
</w:styles>"""

    def escape(text):
        return (text.replace("&", "&amp;")
                    .replace("<", "&lt;")
                    .replace(">", "&gt;")
                    .replace('"', "&quot;")
                    .replace("'", "&apos;"))

    doc_body = []
    
    # Title
    doc_body.append(f"""
    <w:p>
      <w:pPr>
        <w:pBdr>
          <w:bottom w:val="single" w:sz="18" w:space="8" w:color="2563EB"/>
        </w:pBdr>
        <w:spacing w:before="120" w:after="120"/>
      </w:pPr>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
          <w:b/>
          <w:sz w:val="48"/>
          <w:color w:val="1E3A8A"/>
        </w:rPr>
        <w:t>{escape(title)}</w:t>
      </w:r>
    </w:p>
    """)
    
    # Subtitle
    if subtitle:
        doc_body.append(f"""
        <w:p>
          <w:pPr>
            <w:spacing w:before="0" w:after="240"/>
          </w:pPr>
          <w:r>
            <w:rPr>
              <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
              <w:i/>
              <w:sz w:val="26"/>
              <w:color w:val="4B5563"/>
            </w:rPr>
            <w:t>{escape(subtitle)}</w:t>
          </w:r>
        </w:p>
        """)

    for section in sections:
        sec_type = section.get('type')
        if sec_type == 'h1':
            doc_body.append(f"""
            <w:p>
              <w:pPr>
                <w:spacing w:before="280" w:after="140"/>
                <w:shd w:val="clear" w:color="auto" w:fill="F1F5F9"/>
              </w:pPr>
              <w:r>
                <w:rPr>
                  <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                  <w:b/>
                  <w:sz w:val="32"/>
                  <w:color w:val="1E40AF"/>
                </w:rPr>
                <w:t> {escape(section['text'])}</w:t>
              </w:r>
            </w:p>
            """)
        elif sec_type == 'h2':
            doc_body.append(f"""
            <w:p>
              <w:pPr>
                <w:spacing w:before="220" w:after="100"/>
              </w:pPr>
              <w:r>
                <w:rPr>
                  <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                  <w:b/>
                  <w:sz w:val="26"/>
                  <w:color w:val="0F766E"/>
                </w:rPr>
                <w:t>{escape(section['text'])}</w:t>
              </w:r>
            </w:p>
            """)
        elif sec_type == 'p':
            bold_prefix = section.get('bold_prefix', '')
            text = section.get('text', '')
            doc_body.append(f"""
            <w:p>
              <w:pPr>
                <w:spacing w:before="60" w:after="120"/>
              </w:pPr>
              {"<w:r><w:rPr><w:b/><w:color w:val='1F2937'/></w:rPr><w:t>" + escape(bold_prefix) + " </w:t></w:r>" if bold_prefix else ""}
              <w:r>
                <w:t>{escape(text)}</w:t>
              </w:r>
            </w:p>
            """)
        elif sec_type == 'bullet':
            doc_body.append(f"""
            <w:p>
              <w:pPr>
                <w:ind w:left="400" w:hanging="240"/>
                <w:spacing w:before="40" w:after="60"/>
              </w:pPr>
              <w:r>
                <w:rPr><w:color w:val="2563EB"/><w:b/></w:rPr>
                <w:t>•  </w:t>
              </w:r>
              {"<w:r><w:rPr><w:b/><w:color w:val='111827'/></w:rPr><w:t>" + escape(section.get('bold_prefix', '')) + " </w:t></w:r>" if section.get('bold_prefix') else ""}
              <w:r>
                <w:t>{escape(section['text'])}</w:t>
              </w:r>
            </w:p>
            """)
        elif sec_type == 'table':
            headers = section.get('headers', [])
            rows = section.get('rows', [])
            
            tbl_xml = """<w:tbl>
              <w:tblPr>
                <w:tblW w:w="9200" w:type="dxa"/>
                <w:tblBorders>
                  <w:top w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>
                  <w:left w:val="none"/>
                  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>
                  <w:right w:val="none"/>
                  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>
                  <w:insideV w:val="none"/>
                </w:tblBorders>
              </w:tblPr>"""
            
            # Header Row
            tbl_xml += "<w:tr><w:trPr><w:tblHeader/></w:trPr>"
            for h in headers:
                tbl_xml += f"""<w:tc>
                  <w:tcPr>
                    <w:shd w:val="clear" w:color="auto" w:fill="1E3A8A"/>
                    <w:tcMar><w:top w:w="120"/><w:bottom w:w="120"/><w:left w:w="140"/><w:right w:w="140"/></w:tcMar>
                  </w:tcPr>
                  <w:p>
                    <w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr>
                    <w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/></w:rPr><w:t>{escape(h)}</w:t></w:r>
                  </w:p>
                </w:tc>"""
            tbl_xml += "</w:tr>"
            
            # Data Rows
            for idx, r in enumerate(rows):
                bg = "F8FAFC" if idx % 2 == 1 else "FFFFFF"
                tbl_xml += "<w:tr>"
                for cell in r:
                    tbl_xml += f"""<w:tc>
                      <w:tcPr>
                        <w:shd w:val="clear" w:color="auto" w:fill="{bg}"/>
                        <w:tcMar><w:top w:w="100"/><w:bottom w:w="100"/><w:left w:w="140"/><w:right w:w="140"/></w:tcMar>
                      </w:tcPr>
                      <w:p>
                        <w:pPr><w:spacing w:before="40" w:after="40"/></w:pPr>
                        <w:r><w:rPr><w:sz w:val="19"/><w:color w:val="334155"/></w:rPr><w:t>{escape(cell)}</w:t></w:r>
                      </w:p>
                    </w:tc>"""
                tbl_xml += "</w:tr>"
            tbl_xml += "</w:tbl>"
            doc_body.append(tbl_xml)

    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    {"".join(doc_body)}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>"""

    with zipfile.ZipFile(filename, 'w', zipfile.ZIP_DEFLATED) as docx:
        docx.writestr('[Content_Types].xml', content_types)
        docx.writestr('_rels/.rels', rels)
        docx.writestr('word/_rels/document.xml.rels', doc_rels)
        docx.writestr('word/styles.xml', styles)
        docx.writestr('word/document.xml', document_xml)

if __name__ == "__main__":
    os.makedirs("/Users/ivan/Desktop/conversor/docs", exist_ok=True)
    
    sections = [
        {"type": "h1", "text": "1. PROPUESTA DEL PROYECTO"},
        {"type": "h2", "text": "1.1 Grupo beneficiado"},
        {"type": "bullet", "bold_prefix": "Beneficiarios directos:", "text": "Personal médico, personal de enfermería y auxiliares administrativos de centros de salud, clínicas comunitarias y puestos de salud de atención primaria."},
        {"type": "bullet", "bold_prefix": "Beneficiarios indirectos:", "text": "Pacientes de la comunidad (quienes reciben atención más oportuna e historiales digitalizados) y autoridades de salud pública (que reciben reportes consolidados y análisis epidemiológicos al día)."},
        
        {"type": "h2", "text": "1.2 Necesidad o problema real identificado"},
        {"type": "p", "bold_prefix": "Contexto:", "text": "En centros de atención comunitaria y consultorios médicos, el registro de consultas externas, recetas y censos se realiza predominantemente a mano en papel. Esto genera tres graves problemas:"},
        {"type": "bullet", "bold_prefix": "Consumo excesivo de tiempo:", "text": "El personal invierte entre 2 y 4 horas al día en transcripción manual de fichas físicas a hojas de cálculo."},
        {"type": "bullet", "bold_prefix": "Errores por caligrafía compleja:", "text": "Dificultad recurrente para descifrar números de identificación (CUI/DNI), dosis, diagnósticos y nombres, provocando errores en el historial."},
        {"type": "bullet", "bold_prefix": "Riesgo de pérdida de información:", "text": "Deterioro físico o extravío de hojas y retrasos significativos en las estadísticas epidemiológicas mensuales."},
        
        {"type": "h2", "text": "1.3 Intervención tecnológica que realizarás"},
        {"type": "p", "bold_prefix": "Solución:", "text": "Despliegue de la plataforma web DocToExcel AI, una herramienta interactiva que integra Inteligencia Artificial Multimodal para digitalización médica:"},
        {"type": "bullet", "bold_prefix": "Visión Artificial con Google Gemini AI:", "text": "Extracción automática y estructurada en JSON de campos tabulares (No., Nombre, CUI, Fecha de Nacimiento, Sexo, Teléfono, Comunidad, Diagnóstico y Tratamiento)."},
        {"type": "bullet", "bold_prefix": "Verificación Lado a Lado (Side-by-Side):", "text": "Visor de alta resolución con zoom, paneo y rotación del documento original enfrentado a una tabla dinámica editable para supervisión humana inmediata (Human-in-the-Loop)."},
        {"type": "bullet", "bold_prefix": "Exportación Directa a Excel (.xlsx):", "text": "Generación instantánea de hojas de cálculo formateadas mediante SheetJS con un solo clic."},
        {"type": "bullet", "bold_prefix": "Arquitectura Ágil:", "text": "Backend modular en Python (Flask) y frontend interactivo y responsivo en HTML5/CSS3/JavaScript."},

        {"type": "h2", "text": "1.4 Objetivo del proyecto"},
        {"type": "p", "bold_prefix": "Objetivo General:", "text": "Desarrollar y validar una plataforma web asistida por Inteligencia Artificial para la digitalización, extracción estructurada y verificación de registros médicos manuscritos a formato Excel (.xlsx), reduciendo en más de un 80% el tiempo de transcripción manual y minimizando los errores de registro en el personal de salud."},

        {"type": "h1", "text": "2. PLANIFICACIÓN DEL PROYECTO"},
        {"type": "h2", "text": "2.1 ¿Qué harás? (Actividades principales)"},
        {"type": "bullet", "bold_prefix": "Fase 1: Diagnóstico y estandarización:", "text": "Recopilación y análisis de los formatos de fichas clínicas y hojas de consulta física utilizadas en la institución."},
        {"type": "bullet", "bold_prefix": "Fase 2: Calibración y optimización técnica:", "text": "Ajuste del System Prompt y parámetros de visión en Gemini AI para el reconocimiento óptimo de terminología médica y caligrafía compleja."},
        {"type": "bullet", "bold_prefix": "Fase 3: Pruebas de estrés y legibilidad:", "text": "Evaluación con documentos bajo diversas condiciones de iluminación, sombras y orientación."},
        {"type": "bullet", "bold_prefix": "Fase 4: Taller de capacitación:", "text": "Sesión práctica con el personal administrativo y sanitario sobre la captura fotográfica adecuada y flujo de corrección en pantalla."},
        {"type": "bullet", "bold_prefix": "Fase 5: Prueba piloto en entorno real:", "text": "Digitalización asistida de lotes de fichas clínicas generadas durante las jornadas de atención médica."},
        {"type": "bullet", "bold_prefix": "Fase 6: Medición y entrega de informe:", "text": "Cuantificación del tiempo ahorrado, tasa de acierto y entrega del reporte final de resultados."},

        {"type": "h2", "text": "2.2 ¿A quién? (Personas beneficiadas)"},
        {"type": "p", "bold_prefix": "Población objetivo:", "text": "De 10 a 15 profesionales de la salud y personal administrativo encargados del registro diario, beneficiando indirectamente a más de 300 pacientes al mes mediante registros exactos y consultas ágiles."},

        {"type": "h2", "text": "2.3 ¿Dónde? (Lugar o institución)"},
        {"type": "p", "bold_prefix": "Institución sede:", "text": "Centro de Salud Comunitario / Puesto de Salud de Atención Primaria (o Clínica / Consultorio Médico Local)."},

        {"type": "h2", "text": "2.4 ¿Cuándo? (Cronograma de ejecución en Septiembre)"},
        {
            "type": "table",
            "headers": ["Período (Septiembre)", "Fase de Trabajo", "Actividades y Entregables"],
            "rows": [
                ["Semana 1 (1 al 7 de sept.)", "Diagnóstico y Preparación", "Análisis de formularios físicos, recolección de muestras anonimizadas y despliegue del entorno."],
                ["Semana 2 (8 al 14 de sept.)", "Optimización y Pruebas Técnicas", "Calibración del modelo Gemini, ajustes en interfaz lado a lado y pruebas de exportación."],
                ["Semana 3 (15 al 21 de sept.)", "Capacitación y Puesta en Marcha", "Taller práctico con el personal e inicio de la digitalización en jornadas reales."],
                ["Semana 4 (22 al 30 de sept.)", "Evaluación y Cierre", "Medición de métricas de impacto, recolección de feedback y entrega del informe final."]
            ]
        },

        {"type": "h2", "text": "2.5 ¿Cómo? (Metodología de intervención)"},
        {"type": "p", "bold_prefix": "Enfoque:", "text": "Se aplicará una metodología ágil e iterativa centrada en el usuario bajo el modelo Human-in-the-Loop:"},
        {"type": "bullet", "bold_prefix": "1. Captura:", "text": "El usuario toma una fotografía de la hoja con su teléfono o escáner y la carga en la aplicación web."},
        {"type": "bullet", "bold_prefix": "2. Inferencia IA:", "text": "El servidor procesa la imagen y solicita a Gemini la conversión de manuscritos a un esquema JSON estructurado."},
        {"type": "bullet", "bold_prefix": "3. Verificación interactiva:", "text": "La interfaz muestra el documento original y la tabla de datos editable en pantalla dividida para que el usuario valide en segundos."},
        {"type": "bullet", "bold_prefix": "4. Exportación:", "text": "Con un clic se descarga el archivo Excel (.xlsx) listo para consolidación estadística."},

        {"type": "h2", "text": "2.6 ¿Con qué? (Recursos tecnológicos y materiales)"},
        {"type": "bullet", "bold_prefix": "Hardware:", "text": "Computadora/Laptop para la estación de trabajo y smartphone para toma de fotografías."},
        {"type": "bullet", "bold_prefix": "Software Backend:", "text": "Python 3.10+, Flask, Flask-CORS, requests, python-dotenv."},
        {"type": "bullet", "bold_prefix": "Modelos de Inteligencia Artificial:", "text": "Google Gemini API (gemini-2.5-flash, gemini-2.0-flash, gemini-1.5-pro)."},
        {"type": "bullet", "bold_prefix": "Frontend y librerías:", "text": "HTML5, CSS3, JavaScript ES6, SheetJS (xlsx.full.min.js), FontAwesome."},
        {"type": "bullet", "bold_prefix": "Materiales de apoyo:", "text": "Guía rápida de uso en PDF y hojas de control de tiempos de digitación."},

        {"type": "h2", "text": "2.7 ¿Qué resultado esperas obtener?"},
        {"type": "bullet", "bold_prefix": "Reducción de tiempos:", "text": "Disminución del 80% al 90% en el tiempo de transcripción (de 15-20 min por hoja a menos de 1 min)."},
        {"type": "bullet", "bold_prefix": "Alta precisión:", "text": "Tasa de exactitud superior al 95% en los datos consolidados gracias a la supervisión humana guiada."},
        {"type": "bullet", "bold_prefix": "Estandarización:", "text": "Base de datos digitalizada, estructurada y sin errores tipográficos para reportes epidemiológicos."},
        {"type": "bullet", "bold_prefix": "Adopción efectiva:", "text": "Aceptación y uso fluido por parte del personal de salud sin requerir conocimientos técnicos avanzados."},

        {"type": "h2", "text": "2.8 ¿Qué principios éticos deberás aplicar?"},
        {"type": "bullet", "bold_prefix": "1. Confidencialidad y Secreto Médico:", "text": "Protección rigurosa de los datos de salud de los pacientes. No almacenamiento persistente no autorizado en servidores de terceros y anonimización obligatoria en fases de prueba."},
        {"type": "bullet", "bold_prefix": "2. Supervisión Humana Obligatoria (Human-in-the-Loop):", "text": "La IA opera exclusivamente como asistente de captura; el profesional humano siempre revisa, valida y autoriza la información final."},
        {"type": "bullet", "bold_prefix": "3. Integridad y Fidelidad de la Información:", "text": "Garantía de fidelidad con el documento físico, evitando sesgos, adiciones o alucinaciones diagnósticas."},
        {"type": "bullet", "bold_prefix": "4. Seguridad y Protección de Accesos:", "text": "Uso de credenciales protegidas en variables de entorno y protocolos seguros."},
        {"type": "bullet", "bold_prefix": "5. Equidad y Accesibilidad Digital:", "text": "Interfaz clara, amigable y accesible para todo el personal, independientemente de su nivel de alfabetización digital."}
    ]

    target_docx = "/Users/ivan/Desktop/conversor/docs/Propuesta_y_Planificacion_DocToExcel_AI.docx"
    create_docx(
        target_docx,
        "DocToExcel AI: Propuesta y Planificación",
        "Sistema Inteligente de Digitalización de Registros Médicos Manuscritos a Excel con IA",
        sections
    )
    print(f"Documento generado exitosamente en: {target_docx}")
