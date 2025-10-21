package com.example.docx.ui

import android.annotation.SuppressLint
import android.graphics.BitmapFactory
import androidx.compose.foundation.Image
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Edit
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.drawWithContent
import androidx.compose.ui.draw.shadow
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.draw.drawBehind
import androidx.compose.ui.graphics.asImageBitmap
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.SpanStyle
import androidx.compose.ui.text.buildAnnotatedString
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.font.FontStyle
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.BaselineShift
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.text.style.TextDecoration
import androidx.compose.ui.text.withStyle
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.em
import androidx.compose.ui.unit.sp
import androidx.core.net.toUri
import androidx.core.graphics.toColorInt
import androidx.documentfile.provider.DocumentFile
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import java.io.InputStream
import org.apache.poi.xwpf.usermodel.*
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*
import java.math.BigInteger

sealed class DocxElement {
    data class Paragraph(
        val runs: List<Run>,
        val alignment: String? = null,
        val spacingBefore: Int? = null,
        val spacingAfter: Int? = null,
        val lineSpacing: Int? = null,
        val lineSpacingRule: String? = null,
        val indentLeft: Int? = null,
        val indentRight: Int? = null,
        val indentFirstLine: Int? = null,
        val indentHanging: Int? = null,
        val numId: String? = null,
        val ilvl: String? = null,
        val isPageBreak: Boolean = false,
        val backgroundColor: String? = null,
        val borders: ParagraphBorders? = null,
        val styleId: String? = null
    ) : DocxElement()

    data class Table(
        val rows: List<TableRow>,
        val borders: TableBorders? = null,
        val width: Int? = null,
        val alignment: String? = null,
        val cellSpacing: Int? = null,
        val indent: Int? = null
    ) : DocxElement()

    data class TableRow(val cells: List<TableCell>, val height: Int? = null, val isHeader: Boolean = false)

    data class TableCell(
        val content: List<DocxElement>,
        val width: Int? = null,
        val backgroundColor: String? = null,
        val borders: CellBorders? = null,
        val gridSpan: Int = 1,
        val verticalAlignment: String? = null,
        val margins: CellMargins? = null
    )

    data class Run(
        val text: String,
        val bold: Boolean = false,
        val italic: Boolean = false,
        val underline: Boolean = false,
        val strikethrough: Boolean = false,
        val fontSize: Int? = null,
        val fontFamily: String? = null,
        val color: String? = null,
        val highlight: String? = null,
        val imageData: ByteArray? = null,
        val superscript: Boolean = false,
        val subscript: Boolean = false,
        val smallCaps: Boolean = false,
        val allCaps: Boolean = false,
        val doubleStrike: Boolean = false
    ) {
        override fun equals(other: Any?): Boolean {
            if (this === other) return true
            if (javaClass != other?.javaClass) return false

            other as Run

            if (bold != other.bold) return false
            if (italic != other.italic) return false
            if (underline != other.underline) return false
            if (strikethrough != other.strikethrough) return false
            if (fontSize != other.fontSize) return false
            if (superscript != other.superscript) return false
            if (subscript != other.subscript) return false
            if (smallCaps != other.smallCaps) return false
            if (allCaps != other.allCaps) return false
            if (doubleStrike != other.doubleStrike) return false
            if (text != other.text) return false
            if (fontFamily != other.fontFamily) return false
            if (color != other.color) return false
            if (highlight != other.highlight) return false
            if (imageData != null || other.imageData != null) {
                if (imageData == null || other.imageData == null) return false
                if (!imageData.contentEquals(other.imageData)) return false
            }

            return true
        }

        override fun hashCode(): Int {
            var result = bold.hashCode()
            result = 31 * result + italic.hashCode()
            result = 31 * result + underline.hashCode()
            result = 31 * result + strikethrough.hashCode()
            result = 31 * result + (fontSize ?: 0)
            result = 31 * result + superscript.hashCode()
            result = 31 * result + subscript.hashCode()
            result = 31 * result + smallCaps.hashCode()
            result = 31 * result + allCaps.hashCode()
            result = 31 * result + doubleStrike.hashCode()
            result = 31 * result + text.hashCode()
            result = 31 * result + (fontFamily?.hashCode() ?: 0)
            result = 31 * result + (color?.hashCode() ?: 0)
            result = 31 * result + (highlight?.hashCode() ?: 0)
            result = 31 * result + (imageData?.contentHashCode() ?: 0)
            return result
        }
    }

    data class TableBorders(
        val top: BorderStyle? = null,
        val bottom: BorderStyle? = null,
        val left: BorderStyle? = null,
        val right: BorderStyle? = null,
        val insideH: BorderStyle? = null,
        val insideV: BorderStyle? = null
    )

    data class CellBorders(
        val top: BorderStyle? = null,
        val bottom: BorderStyle? = null,
        val left: BorderStyle? = null,
        val right: BorderStyle? = null
    )

    data class ParagraphBorders(
        val top: BorderStyle? = null,
        val bottom: BorderStyle? = null,
        val left: BorderStyle? = null,
        val right: BorderStyle? = null
    )

    data class BorderStyle(val width: Int = 4, val color: String? = "000000", val style: String = "single")
    data class CellMargins(val top: Int? = null, val bottom: Int? = null, val left: Int? = null, val right: Int? = null)
}

data class DocxPage(val elements: List<DocxElement>)

class DocxZipParser {
    private var document: XWPFDocument? = null
    private var headerFooterPolicy: XWPFHeaderFooterPolicy? = null
    private var styles: XWPFStyles? = null

    suspend fun parse(inputStream: InputStream): List<DocxElement> = withContext(Dispatchers.IO) {
        val elements = mutableListOf<DocxElement>()
        try {
            document = XWPFDocument(inputStream)
            styles = document?.styles
            headerFooterPolicy = document?.let { XWPFHeaderFooterPolicy(it) }

            document?.let { doc ->
                // Parse headers
                headerFooterPolicy?.let { policy ->
                    policy.firstPageHeader?.let { header ->
                        elements.addAll(parseHeader(header))
                    }
                    policy.defaultHeader?.let { header ->
                        elements.addAll(parseHeader(header))
                    }
                }

                // Parse body elements
                doc.bodyElements.forEach { element ->
                    when (element) {
                        is XWPFParagraph -> {
                            elements.add(parseParagraph(element))
                        }

                        is XWPFTable -> {
                            elements.add(parseTable(element))
                        }
                    }
                }

                // Parse footers
                headerFooterPolicy?.let { policy ->
                    policy.firstPageFooter?.let { footer ->
                        elements.addAll(parseFooter(footer))
                    }
                    policy.defaultFooter?.let { footer ->
                        elements.addAll(parseFooter(footer))
                    }
                }
            }

            elements
        } catch (e: Exception) {
            e.printStackTrace()
            elements
        }
    }

    private fun parseHeader(header: XWPFHeader): List<DocxElement> {
        val elements = mutableListOf<DocxElement>()
        header.bodyElements.forEach { element ->
            when (element) {
                is XWPFParagraph -> elements.add(parseParagraph(element))
                is XWPFTable -> elements.add(parseTable(element))
            }
        }
        return elements
    }

    private fun parseFooter(footer: XWPFFooter): List<DocxElement> {
        val elements = mutableListOf<DocxElement>()
        footer.bodyElements.forEach { element ->
            when (element) {
                is XWPFParagraph -> elements.add(parseParagraph(element))
                is XWPFTable -> elements.add(parseTable(element))
            }
        }
        return elements
    }

    private fun parseParagraph(paragraph: XWPFParagraph): DocxElement.Paragraph {
        val runs = mutableListOf<DocxElement.Run>()

        paragraph.runs.forEach { run ->
            parseRun(run)?.let { runs.add(it) }
        }

        // Only add default run if there's text but no runs
        if (runs.isEmpty() && paragraph.text.isNotEmpty()) {
            runs.add(
                DocxElement.Run(
                    text = paragraph.text,
                    bold = false,
                    italic = false,
                    underline = false,
                    strikethrough = false,
                    fontSize = null, // Let default sizing handle it
                    fontFamily = null,
                    color = null,
                    highlight = null,
                    imageData = null,
                    superscript = false,
                    subscript = false,
                    smallCaps = false,
                    allCaps = false,
                    doubleStrike = false
                )
            )
        }

        // Get paragraph properties
        val alignment = paragraph.alignment?.toString()
        val spacingBefore = paragraph.spacingBefore
        val spacingAfter = paragraph.spacingAfter
        val lineSpacing = paragraph.spacingBetween?.toInt()
        val lineSpacingRule = paragraph.spacingLineRule?.toString()

        val indentLeft = paragraph.indentationLeft
        val indentRight = paragraph.indentationRight
        val indentFirstLine = paragraph.indentationFirstLine
        val indentHanging = paragraph.indentationHanging

        // Numbering information
        val numId = paragraph.numID?.toString()
        val ilvl = paragraph.numIlvl?.toString()

        // Page break
        val isPageBreak = paragraph.isPageBreak

        // Background color
        val backgroundColor = paragraph.ctp?.pPr?.shd?.fill?.toString()

        // Borders
        val borders = paragraph.ctp?.pPr?.pBdr?.let { pBdr ->
            DocxElement.ParagraphBorders(
                top = pBdr.top?.let { parseBorderStyle(it) },
                bottom = pBdr.bottom?.let { parseBorderStyle(it) },
                left = pBdr.left?.let { parseBorderStyle(it) },
                right = pBdr.right?.let { parseBorderStyle(it) }
            )
        }

        // Style ID
        val styleId = paragraph.style

        return DocxElement.Paragraph(
            runs = runs,
            alignment = alignment,
            spacingBefore = spacingBefore,
            spacingAfter = spacingAfter,
            lineSpacing = lineSpacing,
            lineSpacingRule = lineSpacingRule,
            indentLeft = indentLeft,
            indentRight = indentRight,
            indentFirstLine = indentFirstLine,
            indentHanging = indentHanging,
            numId = numId,
            ilvl = ilvl,
            isPageBreak = isPageBreak,
            backgroundColor = backgroundColor,
            borders = borders,
            styleId = styleId
        )
    }

    private fun parseRun(run: XWPFRun): DocxElement.Run? {
        // Get text from the run
        var text = run.text()
        if (text == null || text.isEmpty()) {
            text = run.getText(0)
        }
        if (text == null || text.isEmpty()) {
            val textBuilder = StringBuilder()
            var pos = 0
            while (pos < 10) {
                val t = run.getText(pos)
                if (t != null && t.isNotEmpty()) {
                    textBuilder.append(t)
                    pos++
                } else {
                    break
                }
            }
            if (textBuilder.isNotEmpty()) {
                text = textBuilder.toString()
            }
        }
        if (text == null) text = ""

        // Get formatting properties
        val bold = run.isBold
        val italic = run.isItalic
        val underline = run.underline != UnderlinePatterns.NONE
        val strikethrough = run.isStrikeThrough
        val doubleStrike = run.isDoubleStrikeThrough
        val fontSize = run.fontSize
        val fontFamily = run.fontFamily
        val color = run.color

        val highlight = run.textHighlightColor?.toString()

        var superscript = false
        var subscript = false
        try {
            val vertAlign = run.verticalAlignment
            when (vertAlign?.toString()?.lowercase()) {
                "superscript" -> superscript = true
                "subscript" -> subscript = true
            }
        } catch (e: Exception) {
            // Ignore if verticalAlignment is not available
        }

        val smallCaps = run.isSmallCaps
        val allCaps = run.isCapitalized

        var imageData: ByteArray? = null
        run.embeddedPictures?.forEach { picture ->
            try {
                val pictureData = picture.pictureData
                imageData = pictureData.data
            } catch (e: Exception) {
                e.printStackTrace()
            }
        }

        if (text.isEmpty() && imageData == null) {
            return null
        }

        return DocxElement.Run(
            text = text,
            bold = bold,
            italic = italic,
            underline = underline,
            strikethrough = strikethrough,
            fontSize = fontSize,
            fontFamily = fontFamily,
            color = color,
            highlight = highlight,
            imageData = imageData,
            superscript = superscript,
            subscript = subscript,
            smallCaps = smallCaps,
            allCaps = allCaps,
            doubleStrike = doubleStrike
        )
    }

    private fun parseTable(table: XWPFTable): DocxElement.Table {
        val rows = mutableListOf<DocxElement.TableRow>()

        table.rows.forEachIndexed { index, row ->
            rows.add(parseTableRow(row, index == 0))
        }

        // Get table properties
        val ctTbl = table.ctTbl
        val borders = ctTbl?.tblPr?.tblBorders?.let { tblBorders ->
            DocxElement.TableBorders(
                top = tblBorders.top?.let { parseBorderStyle(it) },
                bottom = tblBorders.bottom?.let { parseBorderStyle(it) },
                left = tblBorders.left?.let { parseBorderStyle(it) },
                right = tblBorders.right?.let { parseBorderStyle(it) },
                insideH = tblBorders.insideH?.let { parseBorderStyle(it) },
                insideV = tblBorders.insideV?.let { parseBorderStyle(it) }
            )
        }

        val width = table.width
        val alignment = table.tableAlignment?.toString()
        val cellSpacing = table.cellMarginTop
        val indent = ctTbl?.tblPr?.tblInd?.w?.toString()?.toInt()

        return DocxElement.Table(
            rows = rows,
            borders = borders,
            width = width,
            alignment = alignment,
            cellSpacing = cellSpacing,
            indent = indent
        )
    }

    private fun parseTableRow(row: XWPFTableRow, isHeader: Boolean): DocxElement.TableRow {
        val cells = mutableListOf<DocxElement.TableCell>()

        row.tableCells.forEach { cell ->
            cells.add(parseTableCell(cell))
        }

        val height = row.height

        return DocxElement.TableRow(
            cells = cells,
            height = height,
            isHeader = isHeader
        )
    }

    private fun parseTableCell(cell: XWPFTableCell): DocxElement.TableCell {
        val content = mutableListOf<DocxElement>()

        // Parse cell content
        cell.bodyElements.forEach { element ->
            when (element) {
                is XWPFParagraph -> content.add(parseParagraph(element))
                is XWPFTable -> content.add(parseTable(element))
            }
        }

        // Get cell properties
        val width = cell.width
        val backgroundColor = cell.color

        val ctTc = cell.ctTc
        val borders = ctTc?.tcPr?.tcBorders?.let { tcBorders ->
            DocxElement.CellBorders(
                top = tcBorders.top?.let { parseBorderStyle(it) },
                bottom = tcBorders.bottom?.let { parseBorderStyle(it) },
                left = tcBorders.left?.let { parseBorderStyle(it) },
                right = tcBorders.right?.let { parseBorderStyle(it) }
            )
        }

        val gridSpan = ctTc?.tcPr?.gridSpan?.`val`?.toInt() ?: 1
        val verticalAlignment = cell.verticalAlignment?.toString()

        val margins = ctTc?.tcPr?.tcMar?.let { tcMar ->
            DocxElement.CellMargins(
                top = tcMar.top?.w?.toString()?.toInt(),
                bottom = tcMar.bottom?.w?.toString()?.toInt(),
                left = tcMar.left?.w?.toString()?.toInt(),
                right = tcMar.right?.w?.toString()?.toInt()
            )
        }

        return DocxElement.TableCell(
            content = content,
            width = width,
            backgroundColor = backgroundColor,
            borders = borders,
            gridSpan = gridSpan,
            verticalAlignment = verticalAlignment,
            margins = margins
        )
    }

    private fun parseBorderStyle(border: CTBorder): DocxElement.BorderStyle {
        val width = border.sz?.toInt() ?: 4
        val color = border.color?.toString() ?: "000000"
        val style = border.`val`?.toString() ?: "single"

        return DocxElement.BorderStyle(
            width = width,
            color = color,
            style = style
        )
    }
}

private fun paginateElements(elements: List<DocxElement>, maxElementsPerPage: Int = 40): List<DocxPage> {
    val pages = mutableListOf<DocxPage>()
    if (elements.isEmpty()) return pages
    var currentPageElements = mutableListOf<DocxElement>()
    for (element in elements) {
        if ((element is DocxElement.Paragraph && element.isPageBreak) || currentPageElements.size >= maxElementsPerPage) {
            if (currentPageElements.isNotEmpty()) {
                pages.add(DocxPage(currentPageElements))
                currentPageElements = mutableListOf()
            }
        }
        currentPageElements.add(element)
    }
    if (currentPageElements.isNotEmpty()) pages.add(DocxPage(currentPageElements))
    return pages
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun ViewDocxScreen(fileUriString: String, onNavigateBack: () -> Unit, onNavigateToEdit: (String) -> Unit) {
    val context = LocalContext.current
    val documentName = remember(fileUriString) {
        DocumentFile.fromSingleUri(context, fileUriString.toUri())?.name ?: "Document"
    }

    var pages by remember { mutableStateOf<List<DocxPage>>(emptyList()) }
    var isLoading by remember { mutableStateOf(true) }
    var error by remember { mutableStateOf<String?>(null) }

    LaunchedEffect(fileUriString) {
        isLoading = true
        error = null
        try {
            context.contentResolver.openInputStream(fileUriString.toUri())?.use {
                val elements = DocxZipParser().parse(it)
                if (elements.isEmpty()) error = "No content found in document"
                else pages = paginateElements(elements)
            } ?: run { error = "Could not open file" }
        } catch (e: Exception) {
            e.printStackTrace()
            error = "Failed to load document: ${e.message}"
        } finally {
            isLoading = false
        }
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = { Text(documentName, maxLines = 1) },
                navigationIcon = {
                    IconButton(onClick = onNavigateBack) {
                        Icon(Icons.AutoMirrored.Filled.ArrowBack, contentDescription = "Back")
                    }
                },
                actions = {
                    IconButton(onClick = { onNavigateToEdit(fileUriString) }) {
                        Icon(Icons.Filled.Edit, contentDescription = "Edit")
                    }
                }
            )
        },
        modifier = Modifier.fillMaxSize()
    ) { paddingValues ->
        Box(modifier = Modifier
            .fillMaxSize()
            .padding(paddingValues)
            .background(color = Color(0xFFF0F0F0))) {
            when {
                isLoading -> CircularProgressIndicator(modifier = Modifier.align(Alignment.Center))
                error != null -> Text(error!!, modifier = Modifier
                    .align(Alignment.Center)
                    .padding(16.dp), color = MaterialTheme.colorScheme.error)
                pages.isNotEmpty() -> DocxContent(pages)
            }
        }
    }
}

@Composable
fun DocxContent(pages: List<DocxPage>) {
    LazyColumn(
        modifier = Modifier
            .fillMaxSize()
            .background(Color(0xFFE0E0E0)),
        contentPadding = PaddingValues(horizontal = 16.dp, vertical = 32.dp),
        verticalArrangement = Arrangement.spacedBy(32.dp)
    ) {
        items(pages) { page ->
            Box(modifier = Modifier.fillMaxWidth(), contentAlignment = Alignment.Center) {
                Surface(
                    modifier = Modifier
                        .width(794.dp)
                        .heightIn(min = 1123.dp)
                        .shadow(
                            8.dp,
                            RoundedCornerShape(4.dp),
                            spotColor = Color.Black.copy(alpha = 0.3f)
                        ),
                    shape = RoundedCornerShape(4.dp),
                    color = Color.White,
                ) {
                    Column(modifier = Modifier.padding(horizontal = 64.dp, vertical = 72.dp), verticalArrangement = Arrangement.spacedBy(0.dp)) {
                        page.elements.forEach { element ->
                            when (element) {
                                is DocxElement.Paragraph -> ParagraphItem(paragraph = element)
                                is DocxElement.Table -> TableItem(table = element)
                            }
                        }
                    }
                }
            }
        }
    }
}

private fun ensureNonNegativeDp(value: Float): Float = maxOf(0f, value)

@SuppressLint("UseKtx")
@Composable
fun ParagraphItem(paragraph: DocxElement.Paragraph) {
    // Convert twips to dp (1 twip = 1/20 point, 1 point ≈ 1.33 dp)
    val spacingBefore = paragraph.spacingBefore?.let { ensureNonNegativeDp(it / 20f * 1.33f).dp } ?: 0.dp
    val spacingAfter = paragraph.spacingAfter?.let { ensureNonNegativeDp(it / 20f * 1.33f).dp } ?: 0.dp
    val indentLeft = paragraph.indentLeft?.let { ensureNonNegativeDp(it / 20f * 1.33f).dp } ?: 0.dp
    val indentRight = paragraph.indentRight?.let { ensureNonNegativeDp(it / 20f * 1.33f).dp } ?: 0.dp
    val indentFirstLine = paragraph.indentFirstLine?.let { ensureNonNegativeDp(it / 20f * 1.33f).dp } ?: 0.dp
    val indentHanging = paragraph.indentHanging?.let { ensureNonNegativeDp(it / 20f * 1.33f).dp } ?: 0.dp

    // Default font size is 11pt in Word (Calibri 11)
    val defaultFontSize = 11.sp

    val isBullet = paragraph.numId != null

    val backgroundColor = paragraph.backgroundColor?.let {
        try {
            if (it.equals("auto", ignoreCase = true) || it.isEmpty()) {
                Color.Transparent
            } else {
                Color(android.graphics.Color.parseColor("#$it"))
            }
        } catch (e: Exception) {
            Color.Transparent
        }
    } ?: Color.Transparent

    val annotatedString = buildAnnotatedString {
        paragraph.runs.forEach { run ->
            withStyle(
                style = SpanStyle(
                    fontWeight = if (run.bold) FontWeight.Bold else FontWeight.Normal,
                    fontStyle = if (run.italic) FontStyle.Italic else FontStyle.Normal,
                    textDecoration = when {
                        run.underline && run.strikethrough -> TextDecoration.combine(
                            listOf(TextDecoration.Underline, TextDecoration.LineThrough)
                        )
                        run.underline -> TextDecoration.Underline
                        run.strikethrough || run.doubleStrike -> TextDecoration.LineThrough
                        else -> TextDecoration.None
                    },
                    // Word font sizes are in half-points, so divide by 2
                    fontSize = run.fontSize?.let { (it / 2f).sp } ?: defaultFontSize,
                    fontFamily = getFontFamily(run.fontFamily),
                    color = run.color?.let {
                        try {
                            Color(android.graphics.Color.parseColor("#$it"))
                        } catch (e: Exception) {
                            Color.Black
                        }
                    } ?: Color.Black,
                    background = run.highlight?.let {
                        when (it.uppercase()) {
                            "YELLOW" -> Color(0xFFFFFF00)
                            "GREEN" -> Color(0xFF00FF00)
                            "CYAN" -> Color(0xFF00FFFF)
                            "MAGENTA" -> Color(0xFFFF00FF)
                            "BLUE" -> Color(0xFF0000FF)
                            "RED" -> Color(0xFFFF0000)
                            "DARK_BLUE" -> Color(0xFF000080)
                            "DARK_CYAN" -> Color(0xFF008080)
                            "DARK_GREEN" -> Color(0xFF008000)
                            "DARK_MAGENTA" -> Color(0xFF800080)
                            "DARK_RED" -> Color(0xFF800000)
                            "DARK_YELLOW" -> Color(0xFF808000)
                            "DARK_GRAY" -> Color(0xFF808080)
                            "LIGHT_GRAY" -> Color(0xFFC0C0C0)
                            "BLACK" -> Color(0xFF000000)
                            else -> Color.Transparent
                        }
                    } ?: Color.Transparent,
                    baselineShift = when {
                        run.superscript -> BaselineShift.Superscript
                        run.subscript -> BaselineShift.Subscript
                        else -> BaselineShift.None
                    }
                )
            ) {
                var text = run.text
                if (run.allCaps) text = text.uppercase()
                if (run.smallCaps) text = text.lowercase()
                append(text)
            }
        }
    }

    Column(
        modifier = Modifier
            .fillMaxWidth()
            .padding(
                top = spacingBefore,
                bottom = spacingAfter,
                start = indentLeft,
                end = indentRight
            )
            .then(
                if (backgroundColor != Color.Transparent) {
                    Modifier.background(backgroundColor)
                } else {
                    Modifier
                }
            )
            .then(
                if (paragraph.borders != null) {
                    Modifier.drawWithContent {
                        drawContent()
                        paragraph.borders.top?.let { border ->
                            val width = (border.width / 8f)
                            val color = border.color?.toComposeColor() ?: Color.Black
                            drawLine(
                                color,
                                start = androidx.compose.ui.geometry.Offset(0f, 0f),
                                end = androidx.compose.ui.geometry.Offset(size.width, 0f),
                                strokeWidth = width
                            )
                        }
                        paragraph.borders.bottom?.let { border ->
                            val width = (border.width / 8f)
                            val color = border.color?.toComposeColor() ?: Color.Black
                            drawLine(
                                color,
                                start = androidx.compose.ui.geometry.Offset(0f, size.height),
                                end = androidx.compose.ui.geometry.Offset(size.width, size.height),
                                strokeWidth = width
                            )
                        }
                        paragraph.borders.left?.let { border ->
                            val width = (border.width / 8f)
                            val color = border.color?.toComposeColor() ?: Color.Black
                            drawLine(
                                color,
                                start = androidx.compose.ui.geometry.Offset(0f, 0f),
                                end = androidx.compose.ui.geometry.Offset(0f, size.height),
                                strokeWidth = width
                            )
                        }
                        paragraph.borders.right?.let { border ->
                            val width = (border.width / 8f)
                            val color = border.color?.toComposeColor() ?: Color.Black
                            drawLine(
                                color,
                                start = androidx.compose.ui.geometry.Offset(size.width, 0f),
                                end = androidx.compose.ui.geometry.Offset(size.width, size.height),
                                strokeWidth = width
                            )
                        }
                    }
                } else {
                    Modifier
                }
            )
    ) {
        if (annotatedString.isNotEmpty()) {
            val textAlign = when (paragraph.alignment?.lowercase()) {
                "left" -> TextAlign.Left
                "center" -> TextAlign.Center
                "right" -> TextAlign.Right
                "both", "justify" -> TextAlign.Justify
                else -> TextAlign.Start
            }

            // Line spacing calculation
            val rawLineHeight = paragraph.lineSpacing?.let { spacing ->
                when (paragraph.lineSpacingRule?.lowercase()) {
                    "auto", "atleast" -> spacing / 240.0 // 240 is Word's default "single" spacing
                    "exact" -> spacing / 20.0 * 1.33f / 11f // Convert twips to em (relative to 11pt default)
                    else -> 1.15
                }
            } ?: 1.15
            val lineHeight = rawLineHeight.coerceAtLeast(0.5).em

            Row(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(start = if (isBullet) 0.dp else indentFirstLine)
            ) {
                if (isBullet) {
                    val bulletIndent = when (paragraph.ilvl?.toIntOrNull() ?: 0) {
                        0 -> 0.dp
                        1 -> 24.dp
                        2 -> 48.dp
                        else -> 72.dp
                    }
                    Spacer(modifier = Modifier.width(bulletIndent))
                    Text(
                        "•",
                        modifier = Modifier.padding(end = 8.dp),
                        fontSize = defaultFontSize,
                        color = Color.Black
                    )
                }

                Text(
                    text = annotatedString,
                    modifier = Modifier.weight(1f),
                    textAlign = textAlign,
                    lineHeight = lineHeight
                )
            }
        }

        paragraph.runs.forEach { run ->
            if (run.imageData != null) {
                val bitmap = BitmapFactory.decodeByteArray(run.imageData, 0, run.imageData.size)
                if (bitmap != null) {
                    Image(
                        bitmap = bitmap.asImageBitmap(),
                        contentDescription = "Embedded image",
                        modifier = Modifier
                            .fillMaxWidth()
                            .padding(vertical = 8.dp)
                    )
                }
            }
        }

        if (paragraph.isPageBreak) {
            HorizontalDivider(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(vertical = 16.dp),
                thickness = 2.dp,
                color = Color(0xFFCCCCCC)
            )
        }
    }
}

@Composable
fun TableItem(table: DocxElement.Table) {
    val tableIndent = table.indent?.let { ensureNonNegativeDp(it / 20f * 1.33f).dp } ?: 0.dp

    Column(
        modifier = Modifier
            .fillMaxWidth()
            .padding(
                start = tableIndent,
                top = 8.dp,
                bottom = 8.dp
            )
    ) {
        table.rows.forEach { row ->
            val rowHeight = row.height?.let { height ->
                ensureNonNegativeDp(height / 20f * 1.33f).dp
            }

            Row(
                modifier = Modifier
                    .fillMaxWidth()
                    .then(
                        if (rowHeight != null) {
                            Modifier.height(rowHeight)
                        } else {
                            Modifier.height(IntrinsicSize.Min)
                        }
                    ),
                verticalAlignment = Alignment.Top
            ) {
                row.cells.forEach { cell ->
                    val cellWidth = cell.width?.let { width ->
                        ensureNonNegativeDp(width / 20f * 1.33f).dp
                    }
                    val weight = cell.gridSpan.toFloat()

                    // Calculate cell background color
                    val bgColor = cell.backgroundColor?.let {
                        try {
                            if (it.equals("auto", ignoreCase = true) || it.isEmpty()) {
                                Color.Transparent
                            } else {
                                Color(android.graphics.Color.parseColor("#$it"))
                            }
                        } catch (e: Exception) {
                            Color.Transparent
                        }
                    } ?: Color.Transparent

                    // Calculate cell margins with safe defaults (Word default is 0.08" ≈ 5.76 pt)
                    val topMargin = cell.margins?.top?.let { ensureNonNegativeDp(it / 20f * 1.33f).dp } ?: 4.dp
                    val bottomMargin = cell.margins?.bottom?.let { ensureNonNegativeDp(it / 20f * 1.33f).dp } ?: 4.dp
                    val leftMargin = cell.margins?.left?.let { ensureNonNegativeDp(it / 20f * 1.33f).dp } ?: 5.dp
                    val rightMargin = cell.margins?.right?.let { ensureNonNegativeDp(it / 20f * 1.33f).dp } ?: 5.dp

                    Box(
                        modifier = Modifier
                            .then(
                                if (cellWidth != null) {
                                    Modifier.width(cellWidth)
                                } else {
                                    Modifier.weight(weight, fill = true)
                                }
                            )
                            .fillMaxHeight()
                            .background(bgColor)
                            .drawBehind {
                                // Draw borders if they exist
                                cell.borders?.let { borders ->
                                    borders.top?.let { border ->
                                        val width = (border.width / 8f)
                                        val color = border.color?.toComposeColor() ?: Color.Black
                                        drawLine(
                                            color,
                                            start = androidx.compose.ui.geometry.Offset(0f, 0f),
                                            end = androidx.compose.ui.geometry.Offset(
                                                size.width,
                                                0f
                                            ),
                                            strokeWidth = width
                                        )
                                    }
                                    borders.bottom?.let { border ->
                                        val width = (border.width / 8f)
                                        val color = border.color?.toComposeColor() ?: Color.Black
                                        drawLine(
                                            color,
                                            start = androidx.compose.ui.geometry.Offset(
                                                0f,
                                                size.height
                                            ),
                                            end = androidx.compose.ui.geometry.Offset(
                                                size.width,
                                                size.height
                                            ),
                                            strokeWidth = width
                                        )
                                    }
                                    borders.left?.let { border ->
                                        val width = (border.width / 8f)
                                        val color = border.color?.toComposeColor() ?: Color.Black
                                        drawLine(
                                            color,
                                            start = androidx.compose.ui.geometry.Offset(0f, 0f),
                                            end = androidx.compose.ui.geometry.Offset(
                                                0f,
                                                size.height
                                            ),
                                            strokeWidth = width
                                        )
                                    }
                                    borders.right?.let { border ->
                                        val width = (border.width / 8f)
                                        val color = border.color?.toComposeColor() ?: Color.Black
                                        drawLine(
                                            color,
                                            start = androidx.compose.ui.geometry.Offset(
                                                size.width,
                                                0f
                                            ),
                                            end = androidx.compose.ui.geometry.Offset(
                                                size.width,
                                                size.height
                                            ),
                                            strokeWidth = width
                                        )
                                    }
                                }

                                // Draw table borders if cell borders don't exist
                                if (cell.borders == null) {
                                    table.borders?.let { tableBorders ->
                                        tableBorders.insideV?.let { border ->
                                            val width = (border.width / 8f)
                                            val color =
                                                border.color?.toComposeColor() ?: Color.LightGray
                                            drawLine(
                                                color,
                                                start = androidx.compose.ui.geometry.Offset(
                                                    size.width,
                                                    0f
                                                ),
                                                end = androidx.compose.ui.geometry.Offset(
                                                    size.width,
                                                    size.height
                                                ),
                                                strokeWidth = width
                                            )
                                        }
                                        tableBorders.insideH?.let { border ->
                                            val width = (border.width / 8f)
                                            val color =
                                                border.color?.toComposeColor() ?: Color.LightGray
                                            drawLine(
                                                color,
                                                start = androidx.compose.ui.geometry.Offset(
                                                    0f,
                                                    size.height
                                                ),
                                                end = androidx.compose.ui.geometry.Offset(
                                                    size.width,
                                                    size.height
                                                ),
                                                strokeWidth = width
                                            )
                                        }
                                    }
                                }
                            }
                            .padding(
                                top = topMargin,
                                bottom = bottomMargin,
                                start = leftMargin,
                                end = rightMargin
                            ),
                        contentAlignment = when (cell.verticalAlignment?.lowercase()) {
                            "center" -> Alignment.CenterStart
                            "bottom" -> Alignment.BottomStart
                            else -> Alignment.TopStart
                        }
                    ) {
                        Column(
                            modifier = Modifier.fillMaxWidth()
                        ) {
                            cell.content.forEach { element ->
                                when (element) {
                                    is DocxElement.Paragraph -> {
                                        // Reduce paragraph spacing inside table cells
                                        val cellParagraph = element.copy(
                                            spacingBefore = (element.spacingBefore ?: 0) / 4,
                                            spacingAfter = (element.spacingAfter ?: 0) / 4
                                        )
                                        ParagraphItem(paragraph = cellParagraph)
                                    }
                                    is DocxElement.Table -> TableItem(table = element)
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

fun String.toComposeColor(): Color {
    return try {
        val colorString = if (this.startsWith("#")) this else "#$this"
        Color(colorString.toColorInt())
    } catch (e: Exception) {
        Color.Black
    }
}

private fun getFontFamily(fontName: String?): FontFamily {
    return when (fontName?.lowercase()) {
        "times new roman", "times" -> FontFamily.Serif
        "arial", "helvetica" -> FontFamily.SansSerif
        "courier new", "courier" -> FontFamily.Monospace
        "calibri", "segoe ui", "tahoma", "verdana" -> FontFamily.SansSerif
        "georgia" -> FontFamily.Serif
        else -> FontFamily.SansSerif  // Changed default to SansSerif (matches Word's default)
    }
}