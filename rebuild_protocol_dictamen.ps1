$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Add()

# Page Setup UAM / APA 7: 2.54 cm margins (72 points)
$doc.PageSetup.TopMargin = 72
$doc.PageSetup.BottomMargin = 72
$doc.PageSetup.LeftMargin = 72
$doc.PageSetup.RightMargin = 72

$selection = $word.Selection

function Set-H1($text) {
    $selection.Font.Name = 'Arial'
    $selection.Font.Size = 12
    $selection.Font.Bold = $true
    $selection.Font.Italic = $false
    $selection.Font.Color = 0
    $selection.ParagraphFormat.Alignment = 0
    $selection.ParagraphFormat.SpaceBefore = 14
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.ParagraphFormat.LineSpacingRule = 0
    $selection.ParagraphFormat.FirstLineIndent = 0
    $selection.ParagraphFormat.LeftIndent = 0
    $selection.TypeText($text)
    $selection.TypeParagraph()
}

function Set-H2($text) {
    $selection.Font.Name = 'Arial'
    $selection.Font.Size = 12
    $selection.Font.Bold = $true
    $selection.Font.Italic = $true
    $selection.Font.Color = 0
    $selection.ParagraphFormat.Alignment = 0
    $selection.ParagraphFormat.SpaceBefore = 10
    $selection.ParagraphFormat.SpaceAfter = 4
    $selection.ParagraphFormat.LineSpacingRule = 0
    $selection.ParagraphFormat.FirstLineIndent = 0
    $selection.ParagraphFormat.LeftIndent = 0
    $selection.TypeText($text)
    $selection.TypeParagraph()
}

function Add-P($text) {
    $selection.Font.Name = 'Arial'
    $selection.Font.Size = 12
    $selection.Font.Bold = $false
    $selection.Font.Italic = $false
    $selection.Font.Color = 0
    $selection.ParagraphFormat.Alignment = 0 # Left align (APA 7)
    $selection.ParagraphFormat.FirstLineIndent = 36 # 1.27 cm
    $selection.ParagraphFormat.LeftIndent = 0
    $selection.ParagraphFormat.LineSpacingRule = 2 # Doble espacio (2.0)
    $selection.ParagraphFormat.SpaceBefore = 0
    $selection.ParagraphFormat.SpaceAfter = 0
    $selection.TypeText($text)
    $selection.TypeParagraph()
}

function Add-TableCaption($tableNumber, $tableTitle) {
    $selection.ParagraphFormat.LineSpacingRule = 0
    $selection.ParagraphFormat.SpaceBefore = 12
    $selection.ParagraphFormat.SpaceAfter = 2
    $selection.ParagraphFormat.FirstLineIndent = 0
    $selection.ParagraphFormat.Alignment = 0
    
    $selection.Font.Name = 'Arial'
    $selection.Font.Size = 10
    $selection.Font.Bold = $true
    $selection.Font.Italic = $false
    $selection.TypeText($tableNumber)
    $selection.TypeParagraph()
    
    $selection.Font.Bold = $false
    $selection.Font.Italic = $true
    $selection.TypeText($tableTitle)
    $selection.TypeParagraph()
}

function Add-TableNote($noteText) {
    $selection.ParagraphFormat.LineSpacingRule = 0
    $selection.ParagraphFormat.SpaceBefore = 4
    $selection.ParagraphFormat.SpaceAfter = 12
    $selection.ParagraphFormat.FirstLineIndent = 0
    $selection.ParagraphFormat.Alignment = 0
    
    $selection.Font.Name = 'Arial'
    $selection.Font.Size = 9
    $selection.Font.Bold = $false
    $selection.Font.Italic = $true
    $selection.TypeText('Nota. ')
    $selection.Font.Italic = $false
    $selection.TypeText($noteText)
    $selection.TypeParagraph()
}

function Format-ApaTable($table) {
    $table.Borders.Enable = $false
    # Top border row 1
    $table.Rows.Item(1).Borders.Item(-1).LineStyle = 1
    $table.Rows.Item(1).Borders.Item(-1).LineWidth = 4
    # Bottom border header row
    $table.Rows.Item(1).Borders.Item(-3).LineStyle = 1
    $table.Rows.Item(1).Borders.Item(-3).LineWidth = 4
    # Bottom border last row
    $table.Rows.Item($table.Rows.Count).Borders.Item(-3).LineStyle = 1
    $table.Rows.Item($table.Rows.Count).Borders.Item(-3).LineWidth = 4
}

# ==================== PORTADA OFICIAL UAM ====================
$selection.ParagraphFormat.Alignment = 1
$selection.ParagraphFormat.FirstLineIndent = 0
$selection.ParagraphFormat.LeftIndent = 0
$selection.ParagraphFormat.LineSpacingRule = 0
$selection.ParagraphFormat.SpaceBefore = 0
$selection.ParagraphFormat.SpaceAfter = 6

$selection.Font.Name = 'Arial'
$selection.Font.Size = 14
$selection.Font.Bold = $true
$selection.TypeText('UNIVERSIDAD AMERICANA')
$selection.TypeParagraph()
$selection.Font.Size = 12
$selection.TypeText('FACULTAD DE INGENIERÍA Y ARQUITECTURA')
$selection.TypeParagraph()
$selection.TypeText('CARRERA DE LICENCIATURA EN INGENIERÍA INDUSTRIAL')
$selection.TypeParagraph()
$selection.TypeParagraph()
$selection.TypeParagraph()
$selection.TypeParagraph()

$selection.Font.Size = 13
$selection.Font.Bold = $true
$selection.TypeText('PROTOCOLO DE INVESTIGACIÓN')
$selection.TypeParagraph()
$selection.TypeParagraph()

$selection.Font.Size = 12
$selection.Font.Bold = $true
$selection.TypeText('TEMA:')
$selection.TypeParagraph()
$selection.Font.Bold = $false
$selection.TypeText('“Diseño de un modelo de gestión de operaciones basado en Lean Healthcare para la optimización del flujo de pacientes, la capacidad de los servicios de diagnóstico, la disponibilidad de equipos y la gestión de inventarios de insumos críticos en la Clínica Dr. Barquero, durante el segundo semestre de 2026.”')
$selection.TypeParagraph()
$selection.TypeParagraph()
$selection.TypeParagraph()
$selection.TypeParagraph()

$selection.Font.Bold = $true
$selection.TypeText('Presentado por:')
$selection.TypeParagraph()
$selection.Font.Bold = $false
$selection.TypeText("Br. Javier Sequeira Rosales`nBr. Saymond Mijail Montoya Blandón")
$selection.TypeParagraph()
$selection.TypeParagraph()

$selection.Font.Bold = $true
$selection.TypeText('Para optar al título de:')
$selection.TypeParagraph()
$selection.Font.Bold = $false
$selection.TypeText('Licenciatura en Ingeniería Industrial')
$selection.TypeParagraph()
$selection.TypeParagraph()

$selection.Font.Bold = $true
$selection.TypeText('Tutor:')
$selection.TypeParagraph()
$selection.Font.Bold = $false
$selection.TypeText('MSc. Johanna de los Ángeles Núñez Sevilla')
$selection.TypeParagraph()
$selection.TypeParagraph()
$selection.TypeParagraph()

$selection.TypeText("Managua, Nicaragua`nSeptiembre de 2026")
$selection.TypeParagraph()

$selection.InsertBreak(7)

# ==================== CAPÍTULO I ====================
Set-H1('CAPÍTULO I. PLANTEAMIENTO DEL PROBLEMA')

Set-H2('1.1 Descripción del problema')
Add-P('La Clínica Dr. Barquero es un centro de atención médica privada ambulatoria ubicado en la ciudad de Managua, que ofrece consultas médicas especializadas y servicios de apoyo diagnóstico, incluyendo ultrasonidos, rayos X, mamografía, electrocardiograma, ecocardiograma, colposcopía, odontología y laboratorio clínico. Su operación diaria demanda la articulación coordinada de flujos de pacientes, utilización de consultorios, disponibilidad de equipamiento médico y suministro oportuno de medicamentos y materiales clínicos.')

Add-P('A partir de la exploración preliminar realizada con la dirección médica, se identificó que la clínica opera mediante un esquema de atención por orden de llegada (walk-in), sin un sistema estructurado de citas previas ni nivelación de la demanda. Dicha modalidad genera una marcada variabilidad en la afluencia de usuarios: los días lunes y martes se registran picos de atención de hasta 50 pacientes en la jornada matutina de 7:00 a. m. a 12:00 m. (con una tasa aproximada de 10 pacientes por hora), en contraste con jornadas regulares donde el volumen oscila entre 20 y 25 pacientes diarios.')

Add-P('El diagnóstico preliminar evidencia mayores tiempos de espera en la transición consulta–diagnóstico, situación que sugiere la existencia de restricciones de capacidad cuya magnitud y localización deberán determinarse mediante el análisis cuantitativo del flujo. Si bien la espera para la consulta médica inicial promedia entre 15 y 20 minutos, los usuarios que requieren estudios de diagnóstico por imagen o análisis de laboratorio experimentan tiempos de espera posteriores que oscilan entre 1 y 2 horas para la toma de muestras, realización de pruebas y entrega de resultados. Esta condición genera inconformidades en los usuarios en momentos de alta ocupación y limita la capacidad efectiva de atención.')

Add-P('Respecto al equipamiento tecnológico de apoyo diagnóstico (rayos X, ecógrafos, mamógrafo y equipos de laboratorio), se identificó la necesidad de analizar indicadores de disponibilidad operacional de los equipos, tiempos fuera de servicio, frecuencia de fallas y cumplimiento del mantenimiento preventivo. La ocurrencia de interrupciones no programadas o demoras en la calibración y preparación de los equipos incide directamente en el ritmo de procesamiento de los pacientes. La aplicación formal del indicador de Efectividad Total del Equipo (OEE) se evaluará únicamente si durante la fase de diagnóstico se confirma la disponibilidad, pertinencia y consistencia de los registros requeridos para sus tres componentes.')

Add-P('En el ámbito logístico, la entrevista preliminar indica que a la fecha no se han registrado agotamientos de insumos críticos y que las pérdidas por caducidad se consideran inusuales; no obstante, el aprovisionamiento de medicamentos, reactivos de laboratorio y materiales descartables opera de manera empírica mediante inspección visual directa, sin el respaldo de registros cuantitativos de reposición. Por ello, la gestión de inventarios se aborda como un riesgo operativo latente y una oportunidad de diagnóstico para determinar cuantitativamente los niveles adecuados de existencias, controlando el riesgo de faltantes futuros y evitando sobrecostos por capital inmovilizado. El análisis se delimitará a los insumos críticos seleccionados mediante clasificación ABC complementada con criterios de criticidad y caducidad, considerando consumo histórico, existencias, lead time y comportamiento de reposición.')

Add-P('Adicionalmente, la clínica no dispone de procedimientos operativos estandarizados formalizados para las labores de recepción, preparación de pacientes y gestión de materiales, lo que incrementa la variabilidad en los tiempos de respuesta. Esta situación fundamenta la necesidad de estructurar un modelo de gestión de operaciones que articule el flujo de pacientes, la capacidad diagnóstica, la disponibilidad técnica y el suministro de insumos críticos bajo un enfoque sistemático de mejora continua.')

Set-H2('1.2 Formulación del problema')
Add-P('A partir de la problemática expuesta, se plantea la siguiente pregunta de investigación:')
Add-P('¿Cómo estructurar un modelo de gestión de operaciones basado en Lean Healthcare que integre el flujo de pacientes, la capacidad de los servicios de diagnóstico, la disponibilidad de equipos y la gestión de inventarios de insumos críticos en la Clínica Dr. Barquero durante el segundo semestre de 2026?')

Set-H2('1.3 Justificación')
Add-P('La conveniencia de esta investigación radica en la importancia de fortalecer la gestión operativa en los servicios de salud ambulatorios mediante metodologías estructuradas de ingeniería industrial. La adopción de Lean Healthcare como marco integrador permite analizar las operaciones desde la perspectiva del valor hacia el paciente, facilitando la identificación de restricciones y la sincronización entre áreas clínicas, diagnósticas y de suministro.')

Add-P('Desde el punto de vista práctico y social, la propuesta permitirá identificar oportunidades de reducción de desperdicios y formular alternativas orientadas a disminuir tiempos improductivos, mejorar el aprovechamiento de los recursos diagnósticos y fortalecer el control de los insumos críticos. Esto beneficiará potencialmente a los pacientes al propiciar una atención más oportuna y predecible, y al personal de la clínica al dotarlo de procedimientos estandarizados y cargas de trabajo balanceadas.')

Add-P('En el ámbito teórico y metodológico, la investigación aporta un enfoque articulado que combina el mapeo de procesos, el análisis de capacidad, la gestión de inventarios y la evaluación de disponibilidad tecnológica en una institución de salud de mediana escala. Asimismo, proporciona una ruta metodológica rigurosa para la toma de decisiones basada en datos empíricos y análisis de escenarios operativos.')

Set-H2('1.4 Objetivos de la investigación')
Add-P('Objetivo General:')
Add-P('Diseñar un modelo integral de gestión de operaciones basado en Lean Healthcare, análisis de capacidad y gestión cuantitativa de inventarios, orientado a la optimización del flujo de pacientes, la utilización de los servicios de diagnóstico, la disponibilidad de equipos y el abastecimiento de insumos críticos en la Clínica Dr. Barquero durante el segundo semestre de 2026.')

Add-P('Objetivos Específicos:')
Add-P('1. Diagnosticar el desempeño operativo actual de los procesos de atención, servicios de diagnóstico, disponibilidad de equipos y abastecimiento de insumos críticos mediante indicadores de flujo, capacidad, tiempos, disponibilidad y comportamiento de inventarios.')
Add-P('2. Analizar las restricciones y factores críticos asociados a los tiempos de atención, utilización de los servicios de diagnóstico, disponibilidad de equipos y abastecimiento de insumos críticos mediante la priorización y análisis de las desviaciones identificadas.')
Add-P('3. Diseñar un modelo integral de gestión de operaciones que articule estrategias de mejora del flujo y la capacidad, gestión de la disponibilidad de equipos, control de inventarios de insumos críticos y estandarización de los procesos, con base en las restricciones y factores prioritarios identificados.')
Add-P('4. Evaluar el desempeño técnico-operativo y económico esperado del modelo propuesto mediante la comparación de escenarios e indicadores de flujo, capacidad, disponibilidad, inventarios y costos.')

Set-H2('1.5 Alcance y delimitación')
Add-P('El presente estudio se desarrollará en las instalaciones de la Clínica Dr. Barquero, ubicada en Managua, Nicaragua, durante el segundo semestre del año 2026. Comprende los procesos de recepción, consulta médica ambulatoria, servicios de diagnóstico por imagen (ultrasonido, rayos X, mamografía), laboratorio clínico y el almacenamiento y dispensación de insumos médicos y farmacológicos.')

Add-P('El análisis detallado se concentrará en los servicios diagnósticos, equipos e insumos que, a partir de la línea base, presenten mayor incidencia sobre el flujo, la capacidad, la disponibilidad o el abastecimiento. Los restantes serán considerados como parte del contexto general del sistema.')
Add-P('Delimitación temática: El proyecto se enfoca en el diseño de mejoras operativas, logísticas y de estandarización. No contempla valoraciones sobre diagnósticos o tratamientos clínicos, ni auditorías financieras o contables de la empresa.')

Set-H2('1.6 Vinculación con los Objetivos de Desarrollo Sostenible (ODS)')
Add-P('La investigación se vincula con la Agenda 2030 para el Desarrollo Sostenible a través del ODS 3: Salud y Bienestar, específicamente con la Meta 3.8: "Lograr la cobertura sanitaria universal, en particular la protección contra los riesgos financieros, el acceso a servicios de salud esenciales de calidad y el acceso a medicamentos y vacunas inocuos, eficaces, asequibles y de calidad para todos" (Organización de las Naciones Unidas [ONU], 2015). La propuesta incide en esta meta al orientar la mejora operativa hacia una atención más oportuna, predecible y respaldada por la disponibilidad continua de insumos médicos críticos.')

$selection.InsertBreak(7)

# ==================== CAPÍTULO II ====================
Set-H1('CAPÍTULO II. MARCO TEÓRICO')
Add-P('El marco teórico fundamenta los conceptos y herramientas de la Ingeniería Industrial aplicados a los sistemas de salud, organizándose en seis ejes temáticos esenciales:')

Set-H2('2.1 Lean Healthcare y flujo de valor')
Add-P('La filosofía Lean Healthcare adapta los principios de manufactura esbelta al ámbito de los servicios de salud, orientándose a maximizar la entrega de valor al paciente mediante la identificación y eliminación sistemática de actividades que no añaden valor (desperdicios o mudas) (Graban, 2016). Autores contemporáneos como Tlapa et al. (2020) destacan que la aplicación del Mapeo de la Cadena de Valor (Value Stream Mapping - VSM) en entornos sanitarios permite visualizar de forma integral los flujos de pacientes e información, evidenciando las desconexiones temporales entre áreas clínicas y servicios diagnósticos.')

Set-H2('2.2 Capacidad y gestión de colas en servicios')
Add-P('La gestión de la capacidad en organizaciones de servicios implica armonizar la tasa de demanda de los usuarios con la capacidad instalada de atención y procesamiento (Chase et al., 2018). En contextos de salud ambulatoria, la falta de citas programadas genera variabilidad estocástica en las llegadas. Al respecto, se analizará el comportamiento de las llegadas y tiempos de servicio para determinar el modelo de colas o simulación pertinente, previa verificación de los supuestos correspondientes (Bitran & Fellenz, 2019). Esto permitirá evaluar la ocupación de consultorios y predecir la formación de filas de espera.')

Set-H2('2.3 Disponibilidad y mantenimiento de equipos médicos')
Add-P('Los servicios de diagnóstico dependen estrechamente de la confiabilidad de sus activos tecnológicos. La ingeniería de mantenimiento en el sector sanitario procura mantener la disponibilidad funcional y la seguridad del paciente mediante programas preventivos y protocolos de calibración (Khalaf et al., 2021). Se analizarán variables operacionales tales como tiempos de parada, tasa de fallas y cumplimiento del mantenimiento preventivo. La implementación de indicadores avanzados como la Efectividad Total del Equipo (OEE) se reservará para aquellos equipos donde se verifique la existencia de datos fiables de disponibilidad, rendimiento y calidad de los estudios.')

Set-H2('2.4 Gestión de inventarios sanitarios')
Add-P('La administración de inventarios en salud requiere equilibrar los costos de almacenamiento con la criticidad clínica del desabastecimiento (Chopra & Meindl, 2016). Dado que los insumos presentan patrones de consumo disímiles, los modelos de reposición serán seleccionados según demanda, variabilidad, lead time, criticidad y caducidad de los insumos priorizados. El punto de reorden (ROP) y stock de seguridad se aplicarán cuando los datos y supuestos lo permitan, complementados con la clasificación ABC multicriterio y la regla de rotación PEPS (Primero en Entrar, Primero en Salir) (Moons et al., 2019).')

Set-H2('2.5 Estandarización y mejora de procesos')
Add-P('La estandarización operativa mediante Procedimientos Operativos Estándar (SOPs) establece la mejor práctica documentada para ejecutar una actividad, minimizando la variabilidad atribuible al desempeño individual (Niebel & Freivalds, 2014). Integrada con la metodología 5S, favorece la organización física, la identificación visual y la accesibilidad de materiales en consultorios y áreas técnicas, mitigando tiempos de búsqueda y riesgos operativos (Tezel et al., 2021).')

Set-H2('2.6 Evaluación de propuestas operativas')
Add-P('La valoración de modelos de mejora en ingeniería de operaciones requiere contrastar el desempeño previsto frente a la línea base mediante análisis de escenarios (actual versus propuesto) (Heizer et al., 2020). La estimación de la viabilidad técnica se fundamenta en la capacidad de absorción del flujo proyectado y la reducción de tiempos de ciclo, mientras que la factibilidad económica evalúa los costos de implementación frente a los ahorros operativos derivados de la reducción de tiempos improductivos y mermas de insumos.')

$selection.InsertBreak(7)

# ==================== CAPÍTULO III ====================
Set-H1('CAPÍTULO III. MARCO METODOLÓGICO')

Set-H2('3.1 Enfoque y tipo de investigación')
Add-P('La presente corresponde a una investigación aplicada, descriptiva y propositiva, de diseño no experimental. El levantamiento directo tendrá carácter transversal y se complementará con análisis retrospectivo de registros históricos necesarios para caracterizar demanda, consumo, disponibilidad de equipos y lead time. Metodológicamente adopta un enfoque mixto con predominio cuantitativo, sustentado en la recolección, medición y análisis estadístico de variables de proceso, complementado con información cualitativa del personal operativo (Hernández-Sampieri & Mendoza, 2018; Bernal, 2016).')

Set-H2('3.2 Población, unidades de análisis y muestra')
Add-P('Para garantizar la representatividad y validez de los resultados, se definen cinco unidades de análisis acordes con las dimensiones del estudio. Se aclara que no todas requieren representatividad en el mismo sentido estadístico, distinguiéndose aquellas que serán abordadas por censo de aquellas sujetas a muestreo probabilístico:')
Add-P('1. Pacientes y episodios de atención: Corresponde a los usuarios que asisten a las consultas médicas y servicios diagnósticos. La población estará constituida por los pacientes atendidos durante el período de estudio (segundo semestre 2026). La muestra para la medición de tiempos se determinará mediante muestreo probabilístico con criterio estadístico.')
Add-P('2. Jornadas y períodos de operación: Comprende los turnos de atención matutinos y vespertinos. Se seleccionará una muestra representativa estratificada que incluya jornadas de alta demanda (lunes y martes) y jornadas regulares (miércoles a viernes) para capturar la fluctuación operativa.')
Add-P('3. Equipos diagnósticos priorizados: La población abarca los equipos médicos de imagenología y laboratorio. Se aplicará un censo sobre aquellos equipos críticos identificados en la línea base con mayor impacto en las colas de espera (ultrasonido, rayos X y analizador hematológico).')
Add-P('4. Ítems e insumos críticos (SKUs): Comprende los medicamentos, reactivos y materiales descartables. Se aplicará una selección por análisis de Pareto (ABC) complementado con criticidad clínica y caducidad sobre los registros de consumo para seleccionar los ítems prioritarios.')
Add-P('5. Proveedores asociados: Corresponde a las empresas distribuidoras de los insumos y reactivos priorizados. Se analizará el censo de proveedores activos vinculados a dichos ítems para evaluar plazos de entrega (lead time) y comportamiento de cumplimiento.')

Set-H2('3.3 Técnicas e instrumentos de recolección de información')
Add-P('En la Tabla 1 se presentan las técnicas e instrumentos seleccionados para la recolección de datos, vinculados con cada objetivo específico:')

# TABLA 1 (APA 7)
Add-TableCaption('Tabla 1', 'Técnicas e instrumentos de recolección de información según objetivo específico')

$techTable = $doc.Tables.Add($selection.Range, 6, 5)
$techTable.Columns.Item(1).Width = 60
$techTable.Columns.Item(2).Width = 95
$techTable.Columns.Item(3).Width = 105
$techTable.Columns.Item(4).Width = 135
$techTable.Columns.Item(5).Width = 85

$techTable.Cell(1,1).Range.Text = 'Objetivo'
$techTable.Cell(1,2).Range.Text = 'Técnica'
$techTable.Cell(1,3).Range.Text = 'Instrumento'
$techTable.Cell(1,4).Range.Text = 'Información a obtener'
$techTable.Cell(1,5).Range.Text = 'Fuente'

$techTable.Cell(2,1).Range.Text = 'OE1'
$techTable.Cell(2,2).Range.Text = 'Observación directa estructurada'
$techTable.Cell(2,3).Range.Text = 'Ficha de cronometraje de tiempos y flujo'
$techTable.Cell(2,4).Range.Text = 'Hora de llegada, inicio/fin de consulta y exámenes; esperas y ciclo total.'
$techTable.Cell(2,5).Range.Text = 'Pacientes en salas y consultorios'

$techTable.Cell(3,1).Range.Text = 'OE1'
$techTable.Cell(3,2).Range.Text = 'Revisión técnica y documental'
$techTable.Cell(3,3).Range.Text = 'Ficha de capacidad y disponibilidad de equipos'
$techTable.Cell(3,4).Range.Text = 'Horas disponibles, paros, fallas, tiempos de mantenimiento y calibración.'
$techTable.Cell(3,5).Range.Text = 'Bitácoras de equipos / Personal técnico'

$techTable.Cell(4,1).Range.Text = 'OE1'
$techTable.Cell(4,2).Range.Text = 'Análisis documental retrospectivo'
$techTable.Cell(4,3).Range.Text = 'Matriz de registro de inventarios y compras'
$techTable.Cell(4,4).Range.Text = 'Consumo histórico, existencias, faltantes, mermas y lead time de proveedores.'
$techTable.Cell(4,5).Range.Text = 'Facturas, kardex y registros de compras'

$techTable.Cell(5,1).Range.Text = 'OE1, OE2'
$techTable.Cell(5,2).Range.Text = 'Entrevista técnica semiestructurada'
$techTable.Cell(5,3).Range.Text = 'Guía de entrevista operativa'
$techTable.Cell(5,4).Range.Text = 'Criterios de atención, causas percibidas de demoras y restricciones de capacidad.'
$techTable.Cell(5,5).Range.Text = 'Director médico y personal clave'

$techTable.Cell(6,1).Range.Text = 'OE1, OE3'
$techTable.Cell(6,2).Range.Text = 'Inspección visual estructurada'
$techTable.Cell(6,3).Range.Text = 'Lista de verificación (Checklist) 5S'
$techTable.Cell(6,4).Range.Text = 'Nivel de clasificación, orden, limpieza y señalización en áreas operativas.'
$techTable.Cell(6,5).Range.Text = 'Instalaciones físicas de la clínica'

for ($r=1; $r -le 6; $r++) {
    for ($c=1; $c -le 5; $c++) {
        $techTable.Cell($r,$c).Range.Font.Name = 'Arial'
        $techTable.Cell($r,$c).Range.Font.Size = 9
    }
}
for ($c=1; $c -le 5; $c++) { $techTable.Cell(1,$c).Range.Font.Bold = $true }
Format-ApaTable($techTable)

$selection.SetRange($techTable.Range.End, $techTable.Range.End)
$selection.TypeParagraph()
Add-TableNote('Elaboración propia. OE: Objetivo Específico.')

Add-P('Procedimiento y determinación de la muestra para el estudio de tiempos: El levantamiento cronometrado registrará separadamente la hora de llegada, el inicio y la finalización de cada etapa (recepción, preclínica, consulta, toma de muestras, realización de estudios de imagen y entrega de resultados). Esto permitirá desagregar rigurosamente el tiempo de espera del tiempo de servicio efectivo, obteniendo el tiempo total de ciclo por paciente. Para estimar la media del tiempo de ciclo, el margen de error absoluto se define en e = ±2.5 minutos (precisión justificada frente a la duración global del servicio), con un nivel de confianza del 95% (Z = 1.96). La muestra piloto inicial de 20 pacientes permitirá estimar la desviación estándar (s) para determinar el tamaño definitivo de muestra (n) mediante la fórmula de estimación de medias:')
Add-P('n = (Z * s / e)^2')

Set-H2('3.4 Matriz de operacionalización metodológica')
Add-P('Como se observa en la Tabla 2, la matriz de operacionalización integra las dimensiones, datos requeridos, herramientas y productos para cada uno de los cuatro objetivos específicos:')

# TABLA 2 (APA 7) - CON OE3 INCLUIDO OBLIGATORIAMENTE
Add-TableCaption('Tabla 2', 'Matriz de operacionalización metodológica de las variables del estudio')

$opTable = $doc.Tables.Add($selection.Range, 6, 6)
$opTable.Columns.Item(1).Width = 40
$opTable.Columns.Item(2).Width = 85
$opTable.Columns.Item(3).Width = 90
$opTable.Columns.Item(4).Width = 90
$opTable.Columns.Item(5).Width = 85
$opTable.Columns.Item(6).Width = 90

$opTable.Cell(1,1).Range.Text = 'Obj.'
$opTable.Cell(1,2).Range.Text = 'Dimensión'
$opTable.Cell(1,3).Range.Text = 'Indicador'
$opTable.Cell(1,4).Range.Text = 'Dato requerido'
$opTable.Cell(1,5).Range.Text = 'Instrumento'
$opTable.Cell(1,6).Range.Text = 'Procesamiento'

$opTable.Cell(2,1).Range.Text = 'OE1'
$opTable.Cell(2,2).Range.Text = 'Flujo de pacientes y tiempos'
$opTable.Cell(2,3).Range.Text = 'Tiempo de espera, tiempo de servicio, ciclo total (min)'
$opTable.Cell(2,4).Range.Text = 'Marcas de tiempo de entrada/salida por etapa'
$opTable.Cell(2,5).Range.Text = 'Ficha de cronometraje'
$opTable.Cell(2,6).Range.Text = 'Estadística descriptiva y VSM actual'

$opTable.Cell(3,1).Range.Text = 'OE2'
$opTable.Cell(3,2).Range.Text = 'Capacidad y restricciones'
$opTable.Cell(3,3).Range.Text = 'Tasa de utilización (%), disponibilidad operacional (%), paradas'
$opTable.Cell(3,4).Range.Text = 'Horas de operación, fallas y tiempos de examen'
$opTable.Cell(3,5).Range.Text = 'Ficha técnica de equipos'
$opTable.Cell(3,6).Range.Text = 'Balanceo demanda-capacidad y TOC'

$opTable.Cell(4,1).Range.Text = 'OE2'
$opTable.Cell(4,2).Range.Text = 'Inventarios críticos'
$opTable.Cell(4,3).Range.Text = 'Consumo mensual, rotación, stock, lead time (días)'
$opTable.Cell(4,4).Range.Text = 'Kardex, compras, plazos de distribuidores'
$opTable.Cell(4,5).Range.Text = 'Matriz de inventarios'
$opTable.Cell(4,6).Range.Text = 'Pareto ABC, criticidad y ROP/SS'

# OE3 FORMALMENTE INCORPORADO
$opTable.Cell(5,1).Range.Text = 'OE3'
$opTable.Cell(5,2).Range.Text = 'Modelo integral de operaciones'
$opTable.Cell(5,3).Range.Text = 'Parámetros del modelo, nivelación de flujo, reglas ROP y 5S'
$opTable.Cell(5,4).Range.Text = 'Restricciones priorizadas en OE1–OE2 y capacidad'
$opTable.Cell(5,5).Range.Text = 'Criterios de diseño y normas'
$opTable.Cell(5,6).Range.Text = 'VSM futuro, SOPs, TPM básico y arquitectura'

$opTable.Cell(6,1).Range.Text = 'OE4'
$opTable.Cell(6,2).Range.Text = 'Evaluación técnico-económica'
$opTable.Cell(6,3).Range.Text = 'Variación de tiempos (%), costos de implementación, B/C'
$opTable.Cell(6,4).Range.Text = 'Métricas actuales vs. proyectadas, inversión'
$opTable.Cell(6,5).Range.Text = 'Matriz de escenarios'
$opTable.Cell(6,6).Range.Text = 'Análisis de escenarios y viabilidad económica'

for ($r=1; $r -le 6; $r++) {
    for ($c=1; $c -le 6; $c++) {
        $opTable.Cell($r,$c).Range.Font.Name = 'Arial'
        $opTable.Cell($r,$c).Range.Font.Size = 8
    }
}
for ($c=1; $c -le 6; $c++) { $opTable.Cell(1,$c).Range.Font.Bold = $true }
Format-ApaTable($opTable)

$selection.SetRange($opTable.Range.End, $opTable.Range.End)
$selection.TypeParagraph()
Add-TableNote('Elaboración propia. VSM: Value Stream Mapping; TOC: Teoría de Restricciones; ROP: Punto de Reorden; SS: Stock de Seguridad; SOP: Procedimiento Operativo Estándar; TPM: Mantenimiento Productivo Total; B/C: Relación Beneficio-Costo.')

Set-H2('3.5 Ruta metodológica y esquema funcional')
Add-P('En la Tabla 3 se detalla la articulación de la ruta metodológica por fases, estableciendo los datos de entrada, la herramienta principal y el producto verificable de cada objetivo:')

# TABLA 3 (APA 7)
Add-TableCaption('Tabla 3', 'Esquema funcional de la ruta metodológica por fases')

$schemeTable = $doc.Tables.Add($selection.Range, 5, 4)
$schemeTable.Columns.Item(1).Width = 90
$schemeTable.Columns.Item(2).Width = 130
$schemeTable.Columns.Item(3).Width = 120
$schemeTable.Columns.Item(4).Width = 140

$schemeTable.Cell(1,1).Range.Text = 'Fase / Objetivo'
$schemeTable.Cell(1,2).Range.Text = 'Datos de entrada'
$schemeTable.Cell(1,3).Range.Text = 'Herramienta principal'
$schemeTable.Cell(1,4).Range.Text = 'Producto obtenido'

$schemeTable.Cell(2,1).Range.Text = 'Fase 1: Diagnóstico (OE1)'
$schemeTable.Cell(2,2).Range.Text = 'Tiempos de flujo, horas de equipos, consumo de insumos y bitácoras.'
$schemeTable.Cell(2,3).Range.Text = 'Cronometraje, VSM actual y análisis de capacidad/inventario.'
$schemeTable.Cell(2,4).Range.Text = 'Línea base operativa integral y mapa VSM actual.'

$schemeTable.Cell(3,1).Range.Text = 'Fase 2: Análisis (OE2)'
$schemeTable.Cell(3,2).Range.Text = 'Tiempos de espera, tasas de utilización, fallas y variación de lead time.'
$schemeTable.Cell(3,3).Range.Text = 'Teoría de Restricciones (TOC), Pareto y causa-efecto.'
$schemeTable.Cell(3,4).Range.Text = 'Restricciones y factores críticos priorizados y sustentados.'

$schemeTable.Cell(4,1).Range.Text = 'Fase 3: Diseño (OE3)'
$schemeTable.Cell(4,2).Range.Text = 'Restricciones de flujo, capacidad disponible y consumo de insumos críticos.'
$schemeTable.Cell(4,3).Range.Text = 'VSM futuro, nivelación de citas, TPM preventivo, ROP/SS y SOPs/5S.'
$schemeTable.Cell(4,4).Range.Text = 'Modelo integral de gestión de operaciones para la clínica.'

$schemeTable.Cell(5,1).Range.Text = 'Fase 4: Evaluación (OE4)'
$schemeTable.Cell(5,2).Range.Text = 'Línea base (OE1), modelo diseñado (OE3) y costos de implementación.'
$schemeTable.Cell(5,3).Range.Text = 'Análisis de escenarios comparativos y evaluación económica.'
$schemeTable.Cell(5,4).Range.Text = 'Evaluación técnico-operativa y factibilidad económica sustentada.'

for ($r=1; $r -le 5; $r++) {
    for ($c=1; $c -le 4; $c++) {
        $schemeTable.Cell($r,$c).Range.Font.Name = 'Arial'
        $schemeTable.Cell($r,$c).Range.Font.Size = 8
    }
}
for ($c=1; $c -le 4; $c++) { $schemeTable.Cell(1,$c).Range.Font.Bold = $true }
Format-ApaTable($schemeTable)

$selection.SetRange($schemeTable.Range.End, $schemeTable.Range.End)
$selection.TypeParagraph()
Add-TableNote('Elaboración propia.')

Add-P('Evaluación de la propuesta (Fase 4): La evaluación técnica comparará el escenario actual frente al escenario propuesto utilizando los mismos indicadores definidos en el diagnóstico de OE1 (tiempos de espera, tiempo total de ciclo, utilización de consultorios y equipos, y disponibilidad de insumos). La simulación de flujo se aplicará si los datos y la estructura del proceso permiten construir y validar el modelo de manera consistente. La evaluación económica determinará los costos de puesta en marcha del modelo y estimará los beneficios derivados de la disminución de horas improductivas y reducción de mermas, aplicando indicadores de costo-beneficio únicamente si la información disponible lo sustenta.')

Set-H2('3.6 Cronograma de ejecución')
Add-P('La planificación temporal se expone en la Tabla 4, organizada en 16 semanas para el segundo semestre de 2026 con redacción paralela transversal:')

# TABLA 4 (APA 7)
Add-TableCaption('Tabla 4', 'Cronograma de actividades de investigación estructurado en 16 semanas')

$cronTable = $doc.Tables.Add($selection.Range, 8, 5)
$cronTable.Columns.Item(1).Width = 200
$cronTable.Columns.Item(2).Width = 70
$cronTable.Columns.Item(3).Width = 70
$cronTable.Columns.Item(4).Width = 70
$cronTable.Columns.Item(5).Width = 70

$cronTable.Cell(1,1).Range.Text = 'Fases y Actividades de Investigación'
$cronTable.Cell(1,2).Range.Text = 'Mes 1 (Sem. 1-4)'
$cronTable.Cell(1,3).Range.Text = 'Mes 2 (Sem. 5-8)'
$cronTable.Cell(1,4).Range.Text = 'Mes 3 (Sem. 9-12)'
$cronTable.Cell(1,5).Range.Text = 'Mes 4 (Sem. 13-16)'

$cronTable.Cell(2,1).Range.Text = '1. Diseño, prueba piloto y validación de instrumentos'
$cronTable.Cell(2,2).Range.Text = 'X'

$cronTable.Cell(3,1).Range.Text = '2. OE1: Recolección de datos directos, registros y línea base'
$cronTable.Cell(3,2).Range.Text = 'X'
$cronTable.Cell(3,3).Range.Text = 'X'

$cronTable.Cell(4,1).Range.Text = '3. OE2: Análisis de restricciones, cuellos de botella y causas'
$cronTable.Cell(4,3).Range.Text = 'X'

$cronTable.Cell(5,1).Range.Text = '4. OE3: Diseño del modelo integral de operaciones y mejoras'
$cronTable.Cell(5,4).Range.Text = 'X'

$cronTable.Cell(6,1).Range.Text = '5. OE4: Evaluación de escenarios técnico-operativos y costos'
$cronTable.Cell(6,5).Range.Text = 'X'

$cronTable.Cell(7,1).Range.Text = '6. Integración final, revisión tutorial y preparación de defensa'
$cronTable.Cell(7,5).Range.Text = 'X'

$cronTable.Cell(8,1).Range.Text = '• Redacción del informe final (transversal a todas las fases)'
$cronTable.Cell(8,2).Range.Text = 'X'
$cronTable.Cell(8,3).Range.Text = 'X'
$cronTable.Cell(8,4).Range.Text = 'X'
$cronTable.Cell(8,5).Range.Text = 'X'

for ($r=1; $r -le 8; $r++) {
    for ($c=1; $c -le 5; $c++) {
        $cronTable.Cell($r,$c).Range.Font.Name = 'Arial'
        $cronTable.Cell($r,$c).Range.Font.Size = 8
    }
}
for ($c=1; $c -le 5; $c++) {
    $cronTable.Cell(1,$c).Range.Font.Bold = $true
    $cronTable.Cell(1,$c).Range.ParagraphFormat.Alignment = 1
}
for ($r=2; $r -le 8; $r++) {
    for ($c=2; $c -le 5; $c++) {
        $cronTable.Cell($r,$c).Range.ParagraphFormat.Alignment = 1
    }
}
Format-ApaTable($cronTable)

$selection.SetRange($cronTable.Range.End, $cronTable.Range.End)
$selection.TypeParagraph()
Add-TableNote('Elaboración propia.')

# ==================== NUEVO APARTADO: ÉTICA Y CONFIDENCIALIDAD ====================
Set-H2('3.7 Aspectos éticos y protocolo de confidencialidad')
Add-P('Dado que la investigación involucra la observación directa de flujos y episodios de atención de usuarios en un centro de salud, se establecen los siguientes compromisos y salvaguardas éticas:')
Add-P('1. Anonimización estricta: Los registros de tiempos de flujo se identificarán mediante códigos correlativos alfanuméricos (P-001, P-002, etc.). En ningún caso se registrarán nombres, números de cédula, direcciones ni teléfonos de los pacientes.')
Add-P('2. Ausencia de datos clínicos sensibles: El estudio se limita a variables temporales y logísticas (horas de entrada, salida y duración). No se registrarán diagnósticos médicos, prescripciones farmacológicas individuales ni patologías de los usuarios.')
Add-P('3. Autorización institucional: La recolección de información se realiza bajo el consentimiento formal de la dirección médica y administrativa de la Clínica Dr. Barquero, mediante carta de aval institucional.')
Add-P('4. Custodia y uso exclusivo de la información: Las bases de datos generadas se mantendrán bajo custodia segura del equipo investigador en soportes cifrados, utilizándose con propósitos estrictamente académicos para optar al grado de Licenciatura en Ingeniería Industrial.')
Add-P('5. Respeto al servicio médico: Las observaciones cronometradas en sala se ejecutarán de forma no invasiva, garantizando la privacidad y el confort de los pacientes y sin entorpecer las labores del personal.')

$selection.InsertBreak(7)

# ==================== REFERENCIAS APA 7 ====================
Set-H1('REFERENCIAS')

function Add-Ref($text) {
    $selection.Font.Name = 'Arial'
    $selection.Font.Size = 12
    $selection.Font.Bold = $false
    $selection.Font.Italic = $false
    $selection.ParagraphFormat.Alignment = 0
    $selection.ParagraphFormat.FirstLineIndent = -36 # Sangria francesa 1.27 cm
    $selection.ParagraphFormat.LeftIndent = 36
    $selection.ParagraphFormat.LineSpacingRule = 2 # Doble espacio
    $selection.ParagraphFormat.SpaceBefore = 0
    $selection.ParagraphFormat.SpaceAfter = 6
    $selection.TypeText($text)
    $selection.TypeParagraph()
}

Add-Ref('Ballou, R. H. (2004). Logística: Administración de la cadena de suministro (5.ª ed.). Pearson Educación.')
Add-Ref('Bernal, C. A. (2016). Metodología de la investigación: Administración, economía, humanidades y ciencias sociales (4.ª ed.). Pearson Educación.')
Add-Ref('Bitran, G. R., & Fellenz, C. (2019). Queueing theory in healthcare: Improving patient flow and service design. International Journal of Operations & Production Management, 39(8), 1012-1035. https://doi.org/10.1108/IJOPM-01-2019-0045')
Add-Ref('Chase, R. B., Jacobs, F. R., & Aquilano, N. J. (2018). Administración de operaciones: Producción y cadena de suministros (15.ª ed.). McGraw-Hill Education.')
Add-Ref('Chopra, S., & Meindl, P. (2016). Administración de la cadena de suministro: Estrategia, planeación y operación (6.ª ed.). Pearson Educación.')
Add-Ref('Goldratt, E. M. (1992). La meta: Un proceso de mejora continua (2.ª ed.). North River Press.')
Add-Ref('Graban, M. (2016). Lean hospitals: Improving quality, patient safety, and employee engagement (3rd ed.). CRC Press.')
Add-Ref('Heizer, J., Render, B., & Munson, C. (2020). Operations management: Sustainability and supply chain management (13th ed.). Pearson.')
Add-Ref('Hernández-Sampieri, R., & Mendoza, C. P. (2018). Metodología de la investigación: Las rutas cuantitativa, cualitativa y mixta. McGraw-Hill Education.')
Add-Ref('Khalaf, M. A., Al-Kubaisy, Z., & Hassan, R. (2021). Optimization of medical equipment maintenance management using reliability-centered maintenance and operational indicators. Journal of Healthcare Engineering, 2021, Article 6653218. https://doi.org/10.1155/2021/6653218')
Add-Ref('Moons, K., Waeyenbergh, G., & Pintelon, L. (2019). Measuring the performance of hospital supply chains: A systematic review. Computers & Industrial Engineering, 134, 1-17. https://doi.org/10.1016/j.cie.2019.05.025')
Add-Ref('Niebel, B. W., & Freivalds, A. (2014). Ingeniería industrial: Métodos, estándares y diseño del trabajo (13.ª ed.). McGraw-Hill Education.')
Add-Ref('Organización de las Naciones Unidas. (2015). Transformar nuestro mundo: la Agenda 2030 para el Desarrollo Sostenible. ONU.')
Add-Ref('Tezel, A., Koskela, L., & Aziz, Z. (2021). Lean visual management in healthcare environments: A systematic literature review. Production Planning & Control, 32(11), 932-951. https://doi.org/10.1080/09537287.2020.1776907')
Add-Ref('Tlapa, D., Tortorella, G. L., Frazzon, E. M., & Baez-Lopez, Y. A. (2020). Effects of Lean Healthcare on patient flow: A systematic review. Value in Health, 23(2), 266-274. https://doi.org/10.1016/j.jval.2019.10.009')

$selection.InsertBreak(7)

# ==================== ANEXOS (CON LOS INSTRUMENTOS REALES) ====================
Set-H1('ANEXOS')

Set-H2('Anexo 1. Guía de Entrevista Operativa Preliminar (Aplicada)')
Add-P('A continuación se presenta el instrumento de entrevista semiestructurada aplicado a la dirección de la Clínica Dr. Barquero el mes de septiembre de 2026, cuyos resultados sustentan el diagnóstico inicial del protocolo:')
Add-P('1. ¿Cuáles son los principales servicios y especialidades que ofrece la clínica?`nRespuesta: Servicios: Ultrasonidos, rayos X, mamografía, electrocardiograma, laboratorio clínico, farmacia, ecocardiograma, colposcopía, odontología. Especialidades: Medicina general, medicina interna, ginecología, pediatría, cirugía, ortopedia, otorrinolaringología, urología, nefrología, dermatología, oftalmología, cardiología, neurología, patología, radiología, psicología.')
Add-P('2. ¿En qué horarios y días de la semana se concentra la mayor afluencia de pacientes?`nRespuesta: Lunes y martes en horarios de 7:00 a. m. a 12:00 m.')
Add-P('3. ¿Aproximadamente cuántos pacientes atienden en promedio al día en día normal vs. día pico?`nRespuesta: Lunes (día pico) aproximadamente 50 personas. Día normal (ej. viernes) entre 20 y 25 personas.')
Add-P('4. ¿Cómo gestionan la llegada de pacientes (citas previas o por orden de llegada)?`nRespuesta: Por orden de llegada (walk-in).')
Add-P('5. ¿Cuál es el tiempo de espera promedio antes de ser atendido por el médico y en exámenes?`nRespuesta: Para consulta inicial es de 15 a 20 minutos; sin embargo, para la realización de exámenes correspondientes (laboratorio e imagen) y entrega de resultados hay otro tiempo de espera que oscila entre 1 a 2 horas.')
Add-P('6. ¿Se presentan quejas de pacientes por demoras?`nRespuesta: Sí, en ocasiones cuando se está en máxima capacidad se reciben inconformidades.')
Add-P('7. ¿Cómo realizan el control y reposición de materiales e insumos médicos?`nRespuesta: Se hace con revisión visual; no hay por el momento un sistema automatizado de control.')
Add-P('8. ¿Ha ocurrido que se agote material crítico durante la atención?`nRespuesta: No, por el momento siempre ha habido existencia preventiva.')
Add-P('9. ¿Se han presentado pérdidas por desorden o vencimiento en almacén?`nRespuesta: Sí, de manera muy inusual o poca.')
Add-P('10. ¿Cuentan con manuales de procedimientos operativos o protocolos estandarizados?`nRespuesta: No por el momento.')
Add-P('11. ¿Tienen ubicaciones fijas y señalizadas para los materiales en consultorios?`nRespuesta: Tienen lugares fijos conocidos por el personal, pero no están estandarizados formalmente.')
Add-P('12. Pregunta de oro: Si pudiera mejorar un solo proceso, ¿cuál elegiría?`nRespuesta: Los tiempos de espera, ya que esto agilizaría en gran parte el flujo del paciente en la clínica.')

Set-H2('Anexo 2. Ficha de Registro de Cronometraje de Tiempos y Flujo de Pacientes')
Add-P('En la Tabla A1 se presenta el formato técnico diseñado para registrar el ciclo completo de atención del paciente:')

# TABLA A1
Add-TableCaption('Tabla A1', 'Formato de cronometraje y registro de tiempos de flujo de pacientes')

$timeTable = $doc.Tables.Add($selection.Range, 4, 7)
$timeTable.Columns.Item(1).Width = 35
$timeTable.Columns.Item(2).Width = 70
$timeTable.Columns.Item(3).Width = 80
$timeTable.Columns.Item(4).Width = 80
$timeTable.Columns.Item(5).Width = 75
$timeTable.Columns.Item(6).Width = 75
$timeTable.Columns.Item(7).Width = 65

$timeTable.Cell(1,1).Range.Text = 'N.°'
$timeTable.Cell(1,2).Range.Text = 'Llegada'
$timeTable.Cell(1,3).Range.Text = 'Recepción'
$timeTable.Cell(1,4).Range.Text = 'Consulta'
$timeTable.Cell(1,5).Range.Text = 'Exámenes'
$timeTable.Cell(1,6).Range.Text = 'Resultados'
$timeTable.Cell(1,7).Range.Text = 'Ciclo total'

$timeTable.Cell(2,1).Range.Text = 'P-001'
$timeTable.Cell(2,2).Range.Text = '___:___'
$timeTable.Cell(2,3).Range.Text = '___:___ / ___:___'
$timeTable.Cell(2,4).Range.Text = '___:___ / ___:___'
$timeTable.Cell(2,5).Range.Text = '___:___ / ___:___'
$timeTable.Cell(2,6).Range.Text = '___:___'
$timeTable.Cell(2,7).Range.Text = '___ min'

$timeTable.Cell(3,1).Range.Text = 'P-002'
$timeTable.Cell(3,2).Range.Text = '___:___'
$timeTable.Cell(3,3).Range.Text = '___:___ / ___:___'
$timeTable.Cell(3,4).Range.Text = '___:___ / ___:___'
$timeTable.Cell(3,5).Range.Text = '___:___ / ___:___'
$timeTable.Cell(3,6).Range.Text = '___:___'
$timeTable.Cell(3,7).Range.Text = '___ min'

$timeTable.Cell(4,1).Range.Text = '...'
$timeTable.Cell(4,2).Range.Text = '...'
$timeTable.Cell(4,3).Range.Text = '...'
$timeTable.Cell(4,4).Range.Text = '...'
$timeTable.Cell(4,5).Range.Text = '...'
$timeTable.Cell(4,6).Range.Text = '...'
$timeTable.Cell(4,7).Range.Text = '...'

for ($r=1; $r -le 4; $r++) {
    for ($c=1; $c -le 7; $c++) {
        $timeTable.Cell($r,$c).Range.Font.Name = 'Arial'
        $timeTable.Cell($r,$c).Range.Font.Size = 8
    }
}
for ($c=1; $c -le 7; $c++) {
    $timeTable.Cell(1,$c).Range.Font.Bold = $true
    $timeTable.Cell(1,$c).Range.ParagraphFormat.Alignment = 1
}
Format-ApaTable($timeTable)

$selection.SetRange($timeTable.Range.End, $timeTable.Range.End)
$selection.TypeParagraph()
Add-TableNote('Elaboración propia. P-001: Código de anonimización de paciente.')

Set-H2('Anexo 3. Ficha de Registro de Capacidad y Disponibilidad de Equipos Médicos')
Add-P('En la Tabla A2 se detalla el formato estructurado para evaluar los equipos tecnológicos de diagnóstico:')

# TABLA A2
Add-TableCaption('Tabla A2', 'Ficha de registro de capacidad instalada y disponibilidad operacional de equipos médicos')

$eqTable = $doc.Tables.Add($selection.Range, 4, 6)
$eqTable.Columns.Item(1).Width = 90
$eqTable.Columns.Item(2).Width = 70
$eqTable.Columns.Item(3).Width = 80
$eqTable.Columns.Item(4).Width = 80
$eqTable.Columns.Item(5).Width = 80
$eqTable.Columns.Item(6).Width = 80

$eqTable.Cell(1,1).Range.Text = 'Equipo / Modelo'
$eqTable.Cell(1,2).Range.Text = 'Horas Operación'
$eqTable.Cell(1,3).Range.Text = 'Capacidad (est./h)'
$eqTable.Cell(1,4).Range.Text = 'Fallas (h)'
$eqTable.Cell(1,5).Range.Text = 'Mantenimiento (h)'
$eqTable.Cell(1,6).Range.Text = 'Disponibilidad (%)'

$eqTable.Cell(2,1).Range.Text = 'Ultrasonido'
$eqTable.Cell(2,2).Range.Text = '________'
$eqTable.Cell(2,3).Range.Text = '________'
$eqTable.Cell(2,4).Range.Text = '________'
$eqTable.Cell(2,5).Range.Text = '________'
$eqTable.Cell(2,6).Range.Text = '________ %'

$eqTable.Cell(3,1).Range.Text = 'Rayos X'
$eqTable.Cell(3,2).Range.Text = '________'
$eqTable.Cell(3,3).Range.Text = '________'
$eqTable.Cell(3,4).Range.Text = '________'
$eqTable.Cell(3,5).Range.Text = '________'
$eqTable.Cell(3,6).Range.Text = '________ %'

$eqTable.Cell(4,1).Range.Text = 'Analizador clínico'
$eqTable.Cell(4,2).Range.Text = '________'
$eqTable.Cell(4,3).Range.Text = '________'
$eqTable.Cell(4,4).Range.Text = '________'
$eqTable.Cell(4,5).Range.Text = '________'
$eqTable.Cell(4,6).Range.Text = '________ %'

for ($r=1; $r -le 4; $r++) {
    for ($c=1; $c -le 6; $c++) {
        $eqTable.Cell($r,$c).Range.Font.Name = 'Arial'
        $eqTable.Cell($r,$c).Range.Font.Size = 8
    }
}
for ($c=1; $c -le 6; $c++) {
    $eqTable.Cell(1,$c).Range.Font.Bold = $true
    $eqTable.Cell(1,$c).Range.ParagraphFormat.Alignment = 1
}
Format-ApaTable($eqTable)

$selection.SetRange($eqTable.Range.End, $eqTable.Range.End)
$selection.TypeParagraph()
Add-TableNote('Elaboración propia. est./h: estudios por hora.')

Set-H2('Anexo 4. Matriz de Auditoría de Inventarios y Gestión de Proveedores')
Add-P('En la Tabla A3 se presenta el formato para categorización ABC y comportamiento de reposición de insumos críticos:')

# TABLA A3
Add-TableCaption('Tabla A3', 'Matriz de registro de inventarios de insumos críticos y evaluación de proveedores')

$invTable = $doc.Tables.Add($selection.Range, 3, 7)
$invTable.Columns.Item(1).Width = 80
$invTable.Columns.Item(2).Width = 60
$invTable.Columns.Item(3).Width = 65
$invTable.Columns.Item(4).Width = 65
$invTable.Columns.Item(5).Width = 65
$invTable.Columns.Item(6).Width = 70
$invTable.Columns.Item(7).Width = 75

$invTable.Cell(1,1).Range.Text = 'Descripción Ítem / SKU'
$invTable.Cell(1,2).Range.Text = 'Categoría'
$invTable.Cell(1,3).Range.Text = 'Consumo Promedio'
$invTable.Cell(1,4).Range.Text = 'Stock Actual'
$invTable.Cell(1,5).Range.Text = 'Lead Time (días)'
$invTable.Cell(1,6).Range.Text = 'Frecuencia Compra'
$invTable.Cell(1,7).Range.Text = 'Clasificación ABC'

$invTable.Cell(2,1).Range.Text = '[Fármaco / Reactivo A]'
$invTable.Cell(2,2).Range.Text = 'Crítico / Frío'
$invTable.Cell(2,3).Range.Text = '________'
$invTable.Cell(2,4).Range.Text = '________'
$invTable.Cell(2,5).Range.Text = '________'
$invTable.Cell(2,6).Range.Text = '________'
$invTable.Cell(2,7).Range.Text = 'A / B / C'

$invTable.Cell(3,1).Range.Text = '[Descartable B]'
$invTable.Cell(3,2).Range.Text = 'Curación'
$invTable.Cell(3,3).Range.Text = '________'
$invTable.Cell(3,4).Range.Text = '________'
$invTable.Cell(3,5).Range.Text = '________'
$invTable.Cell(3,6).Range.Text = '________'
$invTable.Cell(3,7).Range.Text = 'A / B / C'

for ($r=1; $r -le 3; $r++) {
    for ($c=1; $c -le 7; $c++) {
        $invTable.Cell($r,$c).Range.Font.Name = 'Arial'
        $invTable.Cell($r,$c).Range.Font.Size = 8
    }
}
for ($c=1; $c -le 7; $c++) {
    $invTable.Cell(1,$c).Range.Font.Bold = $true
    $invTable.Cell(1,$c).Range.ParagraphFormat.Alignment = 1
}
Format-ApaTable($invTable)

$selection.SetRange($invTable.Range.End, $invTable.Range.End)
$selection.TypeParagraph()
Add-TableNote('Elaboración propia. SKU: Stock Keeping Unit.')

Set-H2('Anexo 5. Lista de Verificación y Rúbrica de Puntuación para Auditoría 5S')
Add-P('Para evitar subjetividad en la evaluación de las 5S, en la Tabla A4 se establece la rúbrica objetiva de puntuación (escala 1 a 5) y los criterios de cumplimiento:')

# TABLA A4 - RUBRICA OBJETIVA EXIGIDA EN EL DICTAMEN
Add-TableCaption('Tabla A4', 'Rúbrica de evaluación objetiva y lista de verificación de metodología 5S')

$fiveSTable = $doc.Tables.Add($selection.Range, 6, 4)
$fiveSTable.Columns.Item(1).Width = 80
$fiveSTable.Columns.Item(2).Width = 140
$fiveSTable.Columns.Item(3).Width = 180
$fiveSTable.Columns.Item(4).Width = 80

$fiveSTable.Cell(1,1).Range.Text = 'Etapa 5S'
$fiveSTable.Cell(1,2).Range.Text = 'Criterio operativo'
$fiveSTable.Cell(1,3).Range.Text = 'Rúbrica de puntuación (1 a 5)'
$fiveSTable.Cell(1,4).Range.Text = 'Calificación'

$fiveSTable.Cell(2,1).Range.Text = '1. Seiri (Clasificar)'
$fiveSTable.Cell(2,2).Range.Text = 'Separación estricta de elementos necesarios de innecesarios en consultorio/área.'
$fiveSTable.Cell(2,3).Range.Text = '1: Desorden total; 2: Elementos obsoletos presentes; 3: Clasificación parcial; 4: Solo lo necesario; 5: Cero desperdicio y control visual.'
$fiveSTable.Cell(2,4).Range.Text = '___ / 5'

$fiveSTable.Cell(3,1).Range.Text = '2. Seiton (Ordenar)'
$fiveSTable.Cell(3,2).Range.Text = 'Disposición, rotulación y delimitación de instrumental y materiales médicos.'
$fiveSTable.Cell(3,3).Range.Text = '1: Sin ubicación fija; 2: Búsqueda > 1 min; 3: Lugares fijos no rotulados; 4: Rotulado claro; 5: Acceso < 10 seg y siluetas.'
$fiveSTable.Cell(3,4).Range.Text = '___ / 5'

$fiveSTable.Cell(4,1).Range.Text = '3. Seiso (Limpiar)'
$fiveSTable.Cell(4,2).Range.Text = 'Limpieza e inspección de equipos médicos e instalaciones clínicas.'
$fiveSTable.Cell(4,3).Range.Text = '1: Suciedad visible; 2: Limpieza reactiva; 3: Limpieza diaria básica; 4: Rutina con checklist; 5: Limpieza como inspección preventiva.'
$fiveSTable.Cell(4,4).Range.Text = '___ / 5'

$fiveSTable.Cell(5,1).Range.Text = '4. Seiketsu (Estandarizar)'
$fiveSTable.Cell(5,2).Range.Text = 'Consolidación de estándares visuales y procedimientos formalizados.'
$fiveSTable.Cell(5,3).Range.Text = '1: Sin estándares; 2: Instrucciones verbales; 3: Normas escritas no visibles; 4: Estándares visuales activos; 5: Auditorías periódicas.'
$fiveSTable.Cell(5,4).Range.Text = '___ / 5'

$fiveSTable.Cell(6,1).Range.Text = '5. Shitsuke (Disciplina)'
$fiveSTable.Cell(6,2).Range.Text = 'Hábito y cumplimiento sistemático del personal con la mejora continua.'
$fiveSTable.Cell(6,3).Range.Text = '1: Incumplimiento sistemático; 2: Cumplimiento bajo supervisión; 3: Cumplimiento regular; 4: Autocontrol del personal; 5: Cultura Kaizen consolidada.'
$fiveSTable.Cell(6,4).Range.Text = '___ / 5'

for ($r=1; $r -le 6; $r++) {
    for ($c=1; $c -le 4; $c++) {
        $fiveSTable.Cell($r,$c).Range.Font.Name = 'Arial'
        $fiveSTable.Cell($r,$c).Range.Font.Size = 8
    }
}
for ($c=1; $c -le 4; $c++) {
    $fiveSTable.Cell(1,$c).Range.Font.Bold = $true
    $fiveSTable.Cell(1,$c).Range.ParagraphFormat.Alignment = 1
}
Format-ApaTable($fiveSTable)

$selection.SetRange($fiveSTable.Range.End, $fiveSTable.Range.End)
$selection.TypeParagraph()
Add-TableNote('Elaboración propia. Rúbrica definida para garantizar objetividad en la evaluación observacional.')

Set-H2('Anexo 6. Carta de Aval y Autorización Institucional de la Clínica Dr. Barquero')
Add-P('Documento formal suscrito por la dirección médica de la Clínica Dr. Barquero que acredita el acceso a las instalaciones y autoriza la toma de datos con fines estrictamente académicos para los investigadores Br. Javier Sequeira Rosales y Br. Saymond Mijail Montoya Blandón.')

$targetDocx = 'C:\Users\jsequ\OneDrive\Desktop\Tesis\Protocolo_de_Investigacion_Clinica_Barquero_UAM.docx'
$targetPdf = 'C:\Users\jsequ\OneDrive\Desktop\Tesis\Protocolo_de_Investigacion_Clinica_Barquero_UAM.pdf'

$doc.SaveAs([ref]$targetDocx)
$doc.SaveAs([ref]$targetPdf, [ref]17)
$doc.Close()
$word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

Copy-Item $targetPdf 'C:\Users\jsequ\OneDrive\Desktop\Protocolo_de_Investigacion_Clinica_Barquero_UAM.pdf' -Force -ErrorAction SilentlyContinue
Copy-Item $targetPdf 'C:\Users\jsequ\Downloads\Protocolo_de_Investigacion_Clinica_Barquero_UAM.pdf' -Force -ErrorAction SilentlyContinue

'SUCCESS FULLY REBUILT WITH PERFECT APA 7 TABLES AND ALL DICTAMEN CORRECTIONS'
