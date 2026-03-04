package com.example.firstkot

import android.net.Uri
import android.os.Bundle
import android.widget.Button
import android.widget.TextView
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import java.util.zip.ZipInputStream
import javax.xml.parsers.DocumentBuilderFactory
import org.w3c.dom.Element

class MainActivity : AppCompatActivity() {

    data class StudentInput(
        val name: String,
        val matricule: String,
        val caMark: Double?,
        val examMark: Double?
    )

    data class StudentGrade(
        val name: String,
        val matricule: String,
        val caMark: Double,
        val examMark: Double,
        val total: Double,
        val grade: String
    )

    private lateinit var importButton: Button
    private lateinit var exportButton: Button
    private lateinit var statusText: TextView
    private lateinit var previewText: TextView

    private var gradedStudents: List<StudentGrade> = emptyList()

    private val importExcelLauncher = registerForActivityResult(
        ActivityResultContracts.OpenDocument()
    ) { uri: Uri? ->
        if (uri == null) {
            showStatus("No file selected.")
            return@registerForActivityResult
        }
        importAndPreviewExcel(uri)
    }

    private val exportExcelLauncher = registerForActivityResult(
        ActivityResultContracts.CreateDocument(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    ) { uri: Uri? ->
        if (uri == null) {
            showStatus("Export cancelled.")
            return@registerForActivityResult
        }
        exportGradesToExcel(uri)
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        importButton = findViewById(R.id.importButton)
        exportButton = findViewById(R.id.exportButton)
        statusText = findViewById(R.id.statusText)
        previewText = findViewById(R.id.previewText)

        importButton.setOnClickListener {
            importExcelLauncher.launch(
                arrayOf(
                    "application/vnd.ms-excel",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            )
        }

        exportButton.setOnClickListener {
            if (gradedStudents.isEmpty()) {
                showStatus("Nothing to export. Import a file first.")
            } else {
                exportExcelLauncher.launch("graded_students.xlsx")
            }
        }
    }

    private fun importAndPreviewExcel(uri: Uri) {
        runCatching {
            val parsedStudents = mutableListOf<StudentInput>()
            val sharedStrings = mutableListOf<String>()

            contentResolver.openInputStream(uri)?.use { inputStream ->
                // First pass: read sharedStrings.xml (if present) and store strings
                ZipInputStream(inputStream).use { zip ->
                    var entry = zip.nextEntry
                    while (entry != null) {
                        if (entry.name == "xl/sharedStrings.xml") {
                            try {
                                val docFactory = DocumentBuilderFactory.newInstance()
                                docFactory.isNamespaceAware = true
                                val doc = docFactory.newDocumentBuilder().parse(zip)
                                val sis = doc.getElementsByTagName("si")
                                for (i in 0 until sis.length) {
                                    val si = sis.item(i) as Element
                                    val tNodes = si.getElementsByTagName("t")
                                    if (tNodes.length > 0) sharedStrings.add(tNodes.item(0).textContent)
                                    else sharedStrings.add("")
                                }
                            } catch (e: Exception) {
                                // ignore shared strings parsing errors
                            }
                            break
                        }
                        entry = zip.nextEntry
                    }
                }
            } ?: throw IllegalStateException("Could not read selected file.")

            // Second pass: read sheet XML and resolve cell values using sharedStrings
            contentResolver.openInputStream(uri)?.use { inputStream ->
                ZipInputStream(inputStream).use { zip ->
                    var entry = zip.nextEntry
                    while (entry != null) {
                        if (entry.name == "xl/worksheets/sheet1.xml") {
                            val docFactory = DocumentBuilderFactory.newInstance()
                            docFactory.isNamespaceAware = true
                            val doc = docFactory.newDocumentBuilder().parse(zip)

                            val rows = doc.getElementsByTagName("row")
                            for (i in 1 until rows.length) {
                                val row = rows.item(i) as Element
                                val cells = row.getElementsByTagName("c")

                                // prepare placeholders for first 4 columns
                                var name = ""
                                var matricule = ""
                                var caMark: Double? = null
                                var examMark: Double? = null

                                for (j in 0 until cells.length) {
                                    val cell = cells.item(j) as Element
                                    val rAttr = cell.getAttribute("r") // e.g., A2, B2
                                    val colLetters = rAttr.replace(Regex("\\d"), "")
                                    val colIndex = columnLettersToIndex(colLetters)

                                    val tAttr = cell.getAttribute("t")
                                    val value = when {
                                        tAttr == "s" -> {
                                            // shared string; <v> contains index
                                            val v = cell.getElementsByTagName("v").item(0)?.textContent
                                            v?.toIntOrNull()?.let { idx -> sharedStrings.getOrNull(idx) } ?: ""
                                        }
                                        tAttr == "inlineStr" -> {
                                            cell.getElementsByTagName("is").item(0)
                                                ?.let { isNode -> (isNode as Element).getElementsByTagName("t").item(0)?.textContent }
                                                ?: ""
                                        }
                                        else -> cell.getElementsByTagName("v").item(0)?.textContent ?: ""
                                    }

                                    when (colIndex) {
                                        0 -> name = value
                                        1 -> matricule = value
                                        2 -> caMark = value.toDoubleOrNull()
                                        3 -> examMark = value.toDoubleOrNull()
                                    }
                                }

                                if (name.isNotBlank() || matricule.isNotBlank()) {
                                    parsedStudents.add(
                                        StudentInput(
                                            name = name.ifBlank { "Unknown Student" },
                                            matricule = matricule.ifBlank { "N/A" },
                                            caMark = caMark,
                                            examMark = examMark
                                        )
                                    )
                                }
                            }
                            break
                        }
                        entry = zip.nextEntry
                    }
                }
            } ?: throw IllegalStateException("Could not read selected file.")

            gradedStudents = parsedStudents.mapNotNull { calculateStudentGrade(it) }
        }.onSuccess {
            if (gradedStudents.isEmpty()) {
                exportButton.isEnabled = false
                showStatus("No valid student rows found in the Excel file.")
                previewText.text = "Preview is empty. Check columns: Name, Matricule, CA, Exam."
                return
            }

            exportButton.isEnabled = true
            val previewLines = buildString {
                appendLine("Preview (Name | Matricule | Total | Grade)")
                appendLine("-----------------------------------------")
                gradedStudents.forEach { student ->
                    appendLine(
                        "${student.name} | ${student.matricule} | " +
                            "${"%.2f".format(student.total)} | ${student.grade}"
                    )
                }
            }
            previewText.text = previewLines
            showStatus("Imported ${gradedStudents.size} students successfully.")
        }.onFailure { error ->
            exportButton.isEnabled = false
            gradedStudents = emptyList()
            showStatus("Import failed: ${error.message ?: "Unknown error"}")
            previewText.text = "Preview will appear here..."
        }
    }

    private fun calculateStudentGrade(student: StudentInput): StudentGrade? {
        val ca = student.caMark?.coerceIn(0.0, 30.0) ?: 0.0
        val exam = student.examMark?.coerceIn(0.0, 70.0) ?: 0.0

        val total = ca + exam
        if (total !in 0.0..100.0) return null

        return StudentGrade(
            name = student.name,
            matricule = student.matricule,
            caMark = ca,
            examMark = exam,
            total = total,
            grade = gradeFromTotal(total)
        )
    }

    private fun gradeFromTotal(total: Double): String {
        val rounded = total.toInt()
        return when (rounded) {
            in 90..100 -> "A"
            in 80..89 -> "B+"
            in 60..79 -> "B"
            in 55..59 -> "C+"
            in 50..54 -> "C"
            in 45..49 -> "D+"
            in 40..44 -> "D"
            else -> "F"
        }
    }

    private fun columnLettersToIndex(letters: String): Int {
        if (letters.isBlank()) return -1
        var result = 0
        letters.uppercase().forEach { ch ->
            if (ch in 'A'..'Z') {
                result = result * 26 + (ch - 'A' + 1)
            }
        }
        return result - 1
    }

    private fun exportGradesToExcel(uri: Uri) {
        runCatching {
            contentResolver.openOutputStream(uri)?.use { outputStream ->
                val zipOut = java.util.zip.ZipOutputStream(outputStream)
                
                zipOut.putNextEntry(java.util.zip.ZipEntry("_rels/.rels"))
                zipOut.write("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""".toByteArray())
                zipOut.closeEntry()

                zipOut.putNextEntry(java.util.zip.ZipEntry("xl/_rels/workbook.xml.rels"))
                zipOut.write("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""".toByteArray())
                zipOut.closeEntry()

                zipOut.putNextEntry(java.util.zip.ZipEntry("xl/workbook.xml"))
                zipOut.write("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheets>
<sheet name="Grades" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
</sheets>
</workbook>""".toByteArray())
                zipOut.closeEntry()

                zipOut.putNextEntry(java.util.zip.ZipEntry("xl/styles.xml"))
                zipOut.write("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="0"/>
<fonts count="1"><font><sz val="11"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>""".toByteArray())
                zipOut.closeEntry()

                zipOut.putNextEntry(java.util.zip.ZipEntry("xl/worksheets/sheet1.xml"))
                val sheetXml = buildString {
                    append("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="inlineStr"><is><t>Student Name</t></is></c><c r="B1" t="inlineStr"><is><t>Student Grade</t></is></c></row>
""")
                    gradedStudents.forEachIndexed { idx, student ->
                        val rowNum = idx + 2
                        append("""<row r="$rowNum"><c r="A$rowNum" t="inlineStr"><is><t>${student.name}</t></is></c><c r="B$rowNum" t="inlineStr"><is><t>${student.grade}</t></is></c></row>
""")
                    }
                    append("</sheetData>\n</worksheet>")
                }
                zipOut.write(sheetXml.toByteArray())
                zipOut.closeEntry()

                zipOut.putNextEntry(java.util.zip.ZipEntry("[Content_Types].xml"))
                zipOut.write("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>""".toByteArray())
                zipOut.closeEntry()

                zipOut.close()
            } ?: throw IllegalStateException("Could not open export destination.")
        }.onSuccess {
            showStatus("Export completed successfully.")
            Toast.makeText(this, "Excel exported.", Toast.LENGTH_SHORT).show()
        }.onFailure { error ->
            showStatus("Export failed: ${error.message ?: "Unknown error"}")
        }
    }

    private fun showStatus(message: String) {
        statusText.text = message
    }
}
