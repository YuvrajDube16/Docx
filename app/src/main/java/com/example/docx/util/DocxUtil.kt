// util/DocxUtil.kt
package com.example.docx.util

import android.content.Context
import android.net.Uri
import android.util.Base64
import android.widget.Toast
import androidx.documentfile.provider.DocumentFile
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import org.apache.commons.text.StringEscapeUtils
import org.apache.poi.xwpf.usermodel.*
import org.jsoup.Jsoup
import org.jsoup.nodes.Element
import org.jsoup.nodes.TextNode
import java.io.ByteArrayOutputStream
import java.io.IOException
import java.util.*

// Existing functions (read/write) remain unchanged
suspend fun readDocxContent(context: Context, uri: Uri): Result<String> = withContext(Dispatchers.IO) {
    try {
        val documentFile = DocumentFile.fromSingleUri(context, uri)
        if (documentFile?.exists() != true) {
            return@withContext Result.failure(Exception("Document not found or inaccessible"))
        }

        context.contentResolver.openInputStream(uri)?.use { inputStream ->
            val document = XWPFDocument(inputStream)
            val textBuilder = StringBuilder()

            try {
                for (paragraph in document.paragraphs) {
                    textBuilder.append(paragraph.text).append("\n")
                }
                Result.success(textBuilder.toString())
            } catch (e: Exception) {
                Result.failure(Exception("Error reading document content: ${e.message}"))
            } finally {
                document.close()
            }
        } ?: Result.failure(Exception("Could not open document"))
    } catch (e: Exception) {
        e.printStackTrace()
        Result.failure(Exception("Error accessing document: ${e.message}"))
    }
}

suspend fun writeHtmlToDocx(context: Context, uri: Uri, html: String): Result<Unit> = withContext(Dispatchers.IO) {
    try {
        val document = XWPFDocument()
        // The HTML from WebView is JSON-encoded, so we need to unescape it.
        val unescapedHtml = StringEscapeUtils.unescapeJson(html).removeSurrounding("\"")
        val body = Jsoup.parse(unescapedHtml).body()

        // Find the page container and process its children
        val pageContainer = body.selectFirst(".page-container")
        val elementsToProcess = pageContainer?.children() ?: body.children()

        for (element in elementsToProcess) {
            if (element.className() == "page") {
                for (child in element.children()) {
                    when (child.tagName().lowercase(Locale.ROOT)) {
                        "p", "h1", "h2", "h3", "h4", "h5", "h6" -> {
                            val paragraph = document.createParagraph()
                            parseElement(child, paragraph)
                        }
                        // Add table handling if needed in the future
                    }
                }
            }
        }

        // Write to a byte array first
        val byteOutput = ByteArrayOutputStream()
        document.write(byteOutput)

        // Then write to the file
        context.contentResolver.openOutputStream(uri, "w")?.use { outputStream ->
            outputStream.write(byteOutput.toByteArray())
            outputStream.flush()
        } ?: throw IOException("Could not open document for writing")

        withContext(Dispatchers.Main) {
            Toast.makeText(context, "Document saved successfully", Toast.LENGTH_SHORT).show()
        }

        Result.success(Unit)
    } catch (e: Exception) {
        e.printStackTrace()
        Result.failure(Exception("Error writing document: ${e.message}"))
    }
}

private fun parseElement(element: Element, paragraph: XWPFParagraph) {
    for (node in element.childNodes()) {
        if (node is TextNode) {
            val run = paragraph.createRun()
            run.setText(node.text())
        } else if (node is Element) {
            val run = paragraph.createRun()
            when (node.tagName().lowercase(Locale.ROOT)) {
                "b", "strong" -> run.isBold = true
                "i", "em" -> run.isItalic = true
                "u" -> run.setUnderline(UnderlinePatterns.SINGLE)
                "s", "strike" -> run.isStrikeThrough = true
            }
            // Recursively parse child elements to apply nested styles
            parseElement(node, paragraph)
        }
    }
}


// --- Moved and updated functions from ViewDocxScreen ---

/**
 * Converts a XWPFDocument to an HTML string with simulated pages.
 */
internal fun convertDocxToHtml(document: XWPFDocument): String {
    val sb = StringBuilder()
    sb.append("<!doctype html><html><head><meta charset=\"utf-8\"/><meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\"/>")
    sb.append(
        """
        <style>
          body { 
            font-family: sans-serif; 
            background-color: #E0E0E0; 
            padding: 2px;
            margin: 0;
          }
          .page {
            background: white;
            border: 1px solid #BDBDBD;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            margin: 16px auto;
            padding: 8px 6px; /* 8px vertical, 6px horizontal */
            max-width: 800px;
            box-sizing: border-box;
          }
          p { margin: 0 0 1em 0; padding: 0; line-height: 1.5; }
          table { border-collapse: collapse; margin: 8px 0; width: 100%; }
          td, th { border: 1px solid #ccc; padding: 6px 8px; vertical-align: top; }
          img { max-width: 100%; height: auto; }
        </style>
        """.trimIndent()
    )
    sb.append("</head><body><div class='page-container'>")

    // Start the first page
    sb.append("<div class='page'>")

    for (elem in document.bodyElements ?: Collections.emptyList<IBodyElement>()) {
        when (elem) {
            is XWPFParagraph -> {
                val p = elem

                // A page break set on a paragraph starts a new page.
                if (p.isPageBreak) {
                    sb.append("</div><div class='page'>")
                }

                val style = p.style ?: ""

                val paragraphStyle = StringBuilder()
                when (p.alignment) {
                    ParagraphAlignment.CENTER -> paragraphStyle.append("text-align:center;")
                    ParagraphAlignment.RIGHT -> paragraphStyle.append("text-align:right;")
                    ParagraphAlignment.BOTH -> paragraphStyle.append("text-align:justify;")
                    else -> paragraphStyle.append("text-align:left;")
                }
                if (p.indentationFirstLine > 0) paragraphStyle.append("text-indent:${p.indentationFirstLine / 20.0}pt;")
                if (p.indentationLeft > 0) paragraphStyle.append("padding-left:${p.indentationLeft / 20.0}pt;")
                if (p.indentationRight > 0) paragraphStyle.append("padding-right:${p.indentationRight / 20.0}pt;")
                if (p.spacingBefore > 0) paragraphStyle.append("padding-top:${p.spacingBefore / 20.0}pt;")
                if (p.spacingAfter > 0) paragraphStyle.append("padding-bottom:${p.spacingAfter / 20.0}pt;")

                val html = if (style.startsWith("Heading", true)) {
                    val level = extractHeadingLevel(style)
                    "<h$level style=\"$paragraphStyle\">${runsToHtml(p.runs)}</h$level>"
                } else {
                    "<p style=\"$paragraphStyle\">${runsToHtml(p.runs)}</p>"
                }
                sb.append(html)
            }
            is XWPFTable -> {
                sb.append("<table>")
                for (row in elem.rows) {
                    sb.append("<tr>")
                    for (cell in row.tableCells) {
                        sb.append("<td>")
                        for (cp in cell.paragraphs) {
                            sb.append(runsToHtml(cp.runs)).append("<br/>")
                        }
                        sb.append("</td>")
                    }
                    sb.append("</tr>")
                }
                sb.append("</table>")
            }
            else -> sb.append("<!-- unsupported element: ${elem.elementType} -->")
        }
    }

    // Close the last page and container
    sb.append("</div></div></body></html>")
    return sb.toString()
}

internal fun runsToHtml(runs: List<XWPFRun>?): String {
    if (runs.isNullOrEmpty()) return ""
    val sb = StringBuilder()

    for (run in runs) {
        // Handle images
        if (run.embeddedPictures.isNotEmpty()) {
            for (pic in run.embeddedPictures) {
                val data = pic.pictureData?.data
                if (data != null) {
                    val base64 = Base64.encodeToString(data, Base64.NO_WRAP)
                    val ext = pic.pictureData.suggestFileExtension()
                    val mime = when (ext.lowercase()) {
                        "png" -> "image/png"
                        "jpg", "jpeg" -> "image/jpeg"
                        "gif" -> "image/gif"
                        else -> "image/*"
                    }
                    sb.append("<img src=\"data:$mime;base64,$base64\"/>")
                }
            }
            continue // Skip text processing for image runs
        }

        // Handle text
        var text = run.text() ?: ""
        if (text.isEmpty()) {
            val texts = run.ctr.tList.mapNotNull { it.stringValue }
            if (texts.isNotEmpty()) text = texts.joinToString("")
        }
        if (text.isEmpty()) continue

        text = htmlEscape(text)

        val openTags = mutableListOf<String>()
        val closeTags = mutableListOf<String>()

        if (run.isBold) { openTags.add("<b>"); closeTags.add(0, "</b>") }
        if (run.isItalic) { openTags.add("<i>"); closeTags.add(0, "</i>") }
        if (run.underline != UnderlinePatterns.NONE) { openTags.add("<u>"); closeTags.add(0, "</u>") }
        if (run.isStrikeThrough) { openTags.add("<s>"); closeTags.add(0, "</s>") }

        val styleAttrs = StringBuilder()
        if (run.fontSize > 0) styleAttrs.append("font-size:${run.fontSize}pt;")
        run.fontFamily?.let { if (it.isNotBlank()) styleAttrs.append("font-family:'$it';") }
        if (styleAttrs.isNotEmpty()) {
            openTags.add("<span style=\"$styleAttrs\">"); closeTags.add(0, "</span>")
        }

        openTags.forEach { sb.append(it) }
        sb.append(text.replace("\n", "<br/>"))
        closeTags.forEach { sb.append(it) }
    }

    return sb.toString()
}

internal fun htmlEscape(s: String): String {
    return s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\"", "&quot;")
        .replace("'", "&#39;")
}

internal fun extractHeadingLevel(styleName: String): Int {
    val regex = Regex("""\d+""")
    val m = regex.find(styleName)
    return m?.value?.toIntOrNull()?.coerceIn(1, 6) ?: 2
}
