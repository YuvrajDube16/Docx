package com.example.docx.ui

import android.annotation.SuppressLint
import android.graphics.Bitmap
import android.graphics.BitmapFactory
import android.net.Uri
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.Image
import androidx.compose.foundation.background
import androidx.compose.foundation.border
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.itemsIndexed
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.shadow
import androidx.compose.ui.focus.FocusState
import androidx.compose.ui.focus.onFocusChanged
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.asImageBitmap
import androidx.compose.ui.graphics.toArgb
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.res.painterResource
import androidx.compose.ui.text.*
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.font.FontStyle
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.TextFieldValue
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.text.style.TextDecoration
import androidx.compose.ui.unit.TextUnit
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.core.graphics.toColorInt
import com.example.docx.R
import org.apache.poi.util.Units
import org.apache.poi.xwpf.usermodel.*
import java.io.ByteArrayOutputStream
import androidx.core.net.toUri

// Helper: Get font family by name
private fun getFontFamily(fontName: String?): FontFamily {
    return when (fontName?.lowercase()) {
        "times new roman" -> FontFamily.Serif
        "arial" -> FontFamily.SansSerif
        "courier new" -> FontFamily.Monospace
        "calibri" -> FontFamily.SansSerif
        else -> FontFamily.Default
    }
}

// Helper: Filter span styles within a selection
private fun getSpanStylesInRange(
    annotatedString: AnnotatedString,
    start: Int,
    end: Int
): List<AnnotatedString.Range<SpanStyle>> {
    return annotatedString.spanStyles.filter { it.start < end && it.end > start }
}


// EditableBodyElement definitions
sealed class EditableBodyElement {
    data class Paragraph(
        val value: TextFieldValue,
        val original: XWPFParagraph,
        val alignment: ParagraphAlignment = ParagraphAlignment.LEFT,
        val spacingBefore: Int = 0,
        val spacingAfter: Int = 0,
        val lineSpacing: Double = 1.0,
        val originalRuns: List<XWPFRun> = emptyList(),
        val paragraphStyle: String? = null
    ) : EditableBodyElement()

    data class Picture(
        val bitmap: Bitmap,
        val originalRun: XWPFRun,
        val originalParagraph: XWPFParagraph
    ) : EditableBodyElement()

    data class Table(
        val original: XWPFTable,
        val cellValues: List<List<TextFieldValue>> = emptyList()
    ) : EditableBodyElement()
}

// Main Composable
@SuppressLint("UseKtx")
@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun EditDocxScreen(
    fileUriString: String,
    onNavigateBack: () -> Unit
) {
    val context = LocalContext.current
    var editableElements by remember { mutableStateOf<List<EditableBodyElement>>(emptyList()) }
    var focusedElementIndex by remember { mutableStateOf(-1) }
    var currentSelection by remember { mutableStateOf<TextRange?>(null) }
    var isLoading by remember { mutableStateOf(true) }
    var documentToSave by remember { mutableStateOf<XWPFDocument?>(null) }

    val createDocumentLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    ) { uri: Uri? ->
        uri?.let {
            documentToSave?.let { doc ->
                try {
                    context.contentResolver.openOutputStream(it)?.use { outputStream ->
                        doc.write(outputStream)
                    }
                    onNavigateBack()
                } catch (e: Exception) {
                    e.printStackTrace()
                } finally {
                    documentToSave = null
                }
            }
        }
    }

    // Load document
    LaunchedEffect(fileUriString) {
        isLoading = true
        val elements = mutableListOf<EditableBodyElement>()
        try {
            context.contentResolver.openInputStream(Uri.parse(fileUriString))?.use { stream ->
                val document = XWPFDocument(stream)
                document.bodyElements.forEach { bodyElement ->
                    when (bodyElement) {
                        is XWPFParagraph -> {
                            if (bodyElement.runs.any { it.embeddedPictures.isNotEmpty() }) {
                                bodyElement.runs.forEach { run ->
                                    run.embeddedPictures.forEach { pic ->
                                        val picData = pic.pictureData
                                        val bitmap = BitmapFactory.decodeByteArray(picData.data, 0, picData.data.size)
                                        elements.add(EditableBodyElement.Picture(bitmap, run, bodyElement))
                                    }
                                }
                            } else {
                                val annotatedString = buildAnnotatedString {
                                    bodyElement.runs.forEach { run ->
                                        val text = run.text() ?: ""
                                        val style = SpanStyle(
                                            fontFamily = getFontFamily(run.fontFamily),
                                            fontWeight = if (run.isBold) FontWeight.Bold else FontWeight.Normal,
                                            fontStyle = if (run.isItalic) FontStyle.Italic else FontStyle.Normal,
                                            textDecoration = if (run.underline != UnderlinePatterns.NONE) TextDecoration.Underline else TextDecoration.None,
                                            fontSize = if (run.fontSize != -1) run.fontSize.sp else 14.sp,
                                            color = try {
                                                if (run.color != null && run.color != "auto") Color(("#${run.color}").toColorInt()) else Color.Unspecified
                                            } catch (e: Exception) {
                                                Color.Unspecified
                                            }
                                        )
                                        withStyle(style) { append(text) }
                                    }
                                }

                                // Safe access to spacingBetween with null check
                                val lineSpacing = try {
                                    val spacing = bodyElement.spacingBetween
                                    if (spacing > 0) spacing / 240.0 else 1.0
                                } catch (e: Exception) {
                                    1.0
                                }

                                elements.add(EditableBodyElement.Paragraph(
                                    value = TextFieldValue(annotatedString),
                                    original = bodyElement,
                                    alignment = bodyElement.alignment ?: ParagraphAlignment.LEFT,
                                    spacingBefore = bodyElement.spacingBefore,
                                    spacingAfter = bodyElement.spacingAfter,
                                    lineSpacing = lineSpacing,
                                    originalRuns = bodyElement.runs.toList(),
                                    paragraphStyle = bodyElement.style
                                ))
                            }
                        }
                        is XWPFTable -> {
                            val cellValues = bodyElement.rows.map { row ->
                                row.tableCells.map { cell -> TextFieldValue(cell.text) }
                            }
                            elements.add(EditableBodyElement.Table(bodyElement, cellValues))
                        }
                    }
                }
            }
            editableElements = elements
        } catch (e: Exception) {
            e.printStackTrace()
        } finally {
            isLoading = false
        }
    }

    // --- STYLE CHANGE HANDLER ---
    val onStyleChange: (SpanStyle, Boolean) -> Unit = { style, enabled ->
        if (focusedElementIndex != -1) {
            val element = editableElements[focusedElementIndex]
            if (element is EditableBodyElement.Paragraph) {
                val selection = element.value.selection
                if (!selection.collapsed) {
                    val annotatedString = element.value.annotatedString
                    val builder = AnnotatedString.Builder(annotatedString.text)
                    // Reapply *all* other styles except those removed
                    annotatedString.spanStyles.forEach { span ->
                        // If toggling off and this style should be removed, then skip it
                        val overlapsSelection = span.start < selection.end && span.end > selection.start
                        val isSameStyle = span.item == style
                        val shouldRemove = enabled.not() && isSameStyle && overlapsSelection
                        if (!shouldRemove) {
                            builder.addStyle(span.item, span.start, span.end)
                        }
                    }
                    // If toggling on, add style to selection
                    if (enabled) {
                        builder.addStyle(style, selection.start, selection.end)
                    }
                    val newValue = TextFieldValue(
                        builder.toAnnotatedString(),
                        selection = element.value.selection,
                        composition = element.value.composition
                    )
                    editableElements = editableElements.mapIndexed { i, el ->
                        if (i == focusedElementIndex) (el as EditableBodyElement.Paragraph).copy(value = newValue) else el
                    }
                }
            }
        }
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = { Text("Edit Document") },
                navigationIcon = {
                    IconButton(onClick = onNavigateBack) {
                        Icon(Icons.AutoMirrored.Filled.ArrowBack, contentDescription = "Back")
                    }
                },
                actions = {
                    Button(onClick = {
                        val newDoc = XWPFDocument()
                        editableElements.forEach { element ->
                            when (element) {
                                is EditableBodyElement.Paragraph -> {
                                    val newParagraph = newDoc.createParagraph()
                                    newParagraph.alignment = element.alignment
                                    newParagraph.spacingBefore = element.spacingBefore
                                    newParagraph.spacingAfter = element.spacingAfter
                                    if (element.lineSpacing != 1.0)
                                        newParagraph.setSpacingBetween(element.lineSpacing * 240, LineSpacingRule.AUTO)
                                    element.paragraphStyle?.let { newParagraph.style = it }

                                    val annotatedString = element.value.annotatedString
                                    val text = annotatedString.text

                                    if (text.isNotEmpty()) {
                                        // Get all boundaries where span starts/ends change
                                        val boundaries = mutableSetOf(0, text.length).apply {
                                            annotatedString.spanStyles.forEach { add(it.start); add(it.end) }
                                        }.filter { it <= text.length }.sorted()
                                        for (i in 0 until boundaries.size - 1) {
                                            val start = boundaries[i]
                                            val end = boundaries[i+1]
                                            if (start >= end) continue
                                            val run = newParagraph.createRun()
                                            run.setText(text.substring(start, end))
                                            getSpanStylesInRange(annotatedString, start, end).forEach { styleRange ->
                                                val spanStyle = styleRange.item
                                                if (spanStyle.fontWeight == FontWeight.Bold) run.isBold = true
                                                if (spanStyle.fontStyle == FontStyle.Italic) run.isItalic = true
                                                if (spanStyle.textDecoration == TextDecoration.Underline) run.setUnderline(UnderlinePatterns.SINGLE)
                                                when (spanStyle.fontFamily) {
                                                    FontFamily.Serif -> run.fontFamily = "Times New Roman"
                                                    FontFamily.SansSerif -> run.fontFamily = "Arial"
                                                    FontFamily.Monospace -> run.fontFamily = "Courier New"
                                                    else -> {}
                                                }
                                                if (spanStyle.fontSize != TextUnit.Unspecified) {
                                                    run.fontSize = spanStyle.fontSize.value.toInt()
                                                }
                                                if (spanStyle.color != Color.Unspecified) run.setColor(String.format("%06X", 0xFFFFFF and spanStyle.color.toArgb()))
                                            }
                                        }
                                    }
                                }
                                is EditableBodyElement.Picture -> {
                                    val paragraph = newDoc.createParagraph()
                                    paragraph.alignment = element.originalParagraph.alignment ?: ParagraphAlignment.CENTER
                                    val run = paragraph.createRun()
                                    val stream = ByteArrayOutputStream().apply {
                                        element.bitmap.compress(Bitmap.CompressFormat.PNG, 100, this)
                                    }
                                    run.addPicture(
                                        stream.toByteArray().inputStream(),
                                        XWPFDocument.PICTURE_TYPE_PNG,
                                        "image.png",
                                        Units.toEMU(200.0), Units.toEMU(150.0)
                                    )
                                }
                                is EditableBodyElement.Table -> {
                                    val table = newDoc.createTable()
                                    element.original.styleID?.let { table.styleID = it }
                                    element.cellValues.forEachIndexed { rowIndex, row ->
                                        val tableRow = if (rowIndex == 0) table.getRow(0) else table.createRow()
                                        row.forEachIndexed { colIndex, cellValue ->
                                            val cell = if (colIndex < tableRow.tableCells.size) tableRow.getCell(colIndex) else tableRow.createCell()
                                            cell.removeParagraph(0)
                                            val para = cell.addParagraph()
                                            para.createRun().setText(cellValue.text)
                                        }
                                    }
                                }
                            }
                        }
                        documentToSave = newDoc
                        val originalFileName = fileUriString.toUri().lastPathSegment?.substringBeforeLast('.') ?: "document"
                        createDocumentLauncher.launch("$originalFileName-copy.docx")
                    }) { Text("Save As...") }
                }
            )
        }
    ) { paddingValues ->
        Box(
            modifier = Modifier.fillMaxSize()
                .padding(paddingValues)
                .background(Color(0xFFDDDDDD))
        ) {
            if (isLoading) {
                CircularProgressIndicator(modifier = Modifier.align(Alignment.Center))
            } else {
                Column(modifier = Modifier.fillMaxSize()) {
                    FormattingToolbar(
                        onImageSelected = {
                            val newElements = editableElements.toMutableList()
                            val dummyParagraph = XWPFDocument().createParagraph()
                            newElements.add(EditableBodyElement.Picture(it, dummyParagraph.createRun(), dummyParagraph))
                            editableElements = newElements
                        },
                        onStyleChange = onStyleChange,
                        currentSelection = currentSelection
                    )
                    Box(
                        modifier = Modifier.fillMaxSize()
                            .background(Color(0xFFDDDDDD))
                            .padding(vertical = 24.dp),
                        contentAlignment = Alignment.TopCenter
                    ) {
                        LazyColumn(
                            modifier = Modifier.width(800.dp)
                                .fillMaxHeight()
                                .shadow(6.dp, RoundedCornerShape(2.dp))
                                .border(1.dp, Color(0xFFB0B0B0))
                                .background(Color.White),
                            contentPadding = PaddingValues(horizontal = 48.dp, vertical = 56.dp),
                            verticalArrangement = Arrangement.spacedBy(8.dp)
                        ) {
                            itemsIndexed(editableElements) { index, element ->
                                when (element) {
                                    is EditableBodyElement.Paragraph -> {
                                        EditableParagraphItem(
                                            paragraph = element,
                                            onValueChange = { newValue ->
                                                currentSelection = newValue.selection
                                                editableElements = editableElements.mapIndexed { i, el ->
                                                    if (i == index)
                                                        (el as EditableBodyElement.Paragraph).copy(value = newValue)
                                                    else el
                                                }
                                            },
                                            onFocusChanged = { isFocused ->
                                                if (isFocused) focusedElementIndex = index
                                            },
                                            modifier = Modifier.fillMaxWidth()
                                        )
                                    }
                                    is EditableBodyElement.Picture -> {
                                        EditablePictureItem(picture = element, modifier = Modifier.fillMaxWidth())
                                    }
                                    is EditableBodyElement.Table -> {
                                        EditableTableItem(
                                            table = element,
                                            onCellValueChange = { rowIndex, colIndex, newValue ->
                                                editableElements = editableElements.mapIndexed { i, el ->
                                                    if (i == index && el is EditableBodyElement.Table) {
                                                        val newCellValues = el.cellValues.mapIndexed { r, row ->
                                                            if (r == rowIndex) {
                                                                row.mapIndexed { c, cell ->
                                                                    if (c == colIndex) newValue else cell
                                                                }
                                                            } else {
                                                                row
                                                            }
                                                        }
                                                        el.copy(cellValues = newCellValues)
                                                    } else el
                                                }
                                            },
                                            modifier = Modifier.fillMaxWidth()
                                        )
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

// Editable components

@Composable
fun EditableParagraphItem(
    paragraph: EditableBodyElement.Paragraph,
    onValueChange: (TextFieldValue) -> Unit,
    onFocusChanged: (Boolean) -> Unit,
    modifier: Modifier = Modifier
) {
    val textAlign = when (paragraph.alignment) {
        ParagraphAlignment.LEFT -> TextAlign.Left
        ParagraphAlignment.CENTER -> TextAlign.Center
        ParagraphAlignment.RIGHT -> TextAlign.Right
        ParagraphAlignment.BOTH -> TextAlign.Justify
        else -> TextAlign.Start
    }
    TextField(
        value = paragraph.value,
        onValueChange = onValueChange,
        textStyle = TextStyle(textAlign = textAlign, lineHeight = 20.sp, fontSize = 14.sp),
        placeholder = { if (paragraph.value.text.isEmpty()) Text("Type here...") },
        modifier = modifier.fillMaxWidth().onFocusChanged { focusState: FocusState -> onFocusChanged(focusState.isFocused) }
    )
}

@Composable
fun EditablePictureItem(picture: EditableBodyElement.Picture, modifier: Modifier = Modifier) {
    val alignment = when (picture.originalParagraph.alignment) {
        ParagraphAlignment.CENTER -> Alignment.Center
        ParagraphAlignment.RIGHT -> Alignment.CenterEnd
        else -> Alignment.CenterStart
    }
    Box(modifier = modifier.padding(vertical = 8.dp), contentAlignment = alignment) {
        Image(bitmap = picture.bitmap.asImageBitmap(), contentDescription = "Document Image", modifier = Modifier.fillMaxWidth())
    }
}

@Composable
fun EditableTableItem(
    table: EditableBodyElement.Table,
    onCellValueChange: (Int, Int, TextFieldValue) -> Unit,
    modifier: Modifier = Modifier
) {
    Column(modifier = modifier.padding(vertical = 8.dp).border(1.dp, Color.Gray)) {
        table.cellValues.forEachIndexed { rowIndex, row ->
            Row(modifier = Modifier.fillMaxWidth().border(0.5.dp, Color.Gray).height(IntrinsicSize.Min)) {
                val totalWidth = try {
                    val widths = table.original.rows.getOrNull(rowIndex)?.tableCells?.mapNotNull { cell ->
                        try {
                            cell.width.takeIf { it > 0 }
                        } catch (e: Exception) {
                            null
                        }
                    } ?: emptyList()
                    widths.sum().takeIf { it > 0 } ?: row.size
                } catch (e: Exception) { row.size }

                row.forEachIndexed { colIndex, cellValue ->
                    val weight = try {
                        val cell = table.original.rows.getOrNull(rowIndex)?.tableCells?.getOrNull(colIndex)
                        val cellWidth = cell?.let {
                            try {
                                it.width.takeIf { w -> w > 0 }
                            } catch (e: Exception) {
                                null
                            }
                        }
                        if (cellWidth != null && cellWidth > 0 && totalWidth > 0) {
                            cellWidth.toFloat() / totalWidth
                        } else {
                            1f / row.size.coerceAtLeast(1)
                        }
                    } catch (e: Exception) {
                        1f / row.size.coerceAtLeast(1)
                    }
                    OutlinedTextField(
                        value = cellValue,
                        onValueChange = { onCellValueChange(rowIndex, colIndex, it) },
                        textStyle = TextStyle(fontSize = 12.sp, fontFamily = FontFamily.Default),
                        placeholder = { if (cellValue.text.isEmpty()) Text("Cell", style = TextStyle(fontSize = 12.sp)) },
                        modifier = Modifier.weight(weight).fillMaxHeight()
                    )
                }
            }
        }
    }
}

@Composable
fun FormattingToolbar(
    onImageSelected: (Bitmap) -> Unit,
    onStyleChange: (SpanStyle, Boolean) -> Unit,
    currentSelection: TextRange?
) {
    var isBold by remember { mutableStateOf(false) }
    var isItalic by remember { mutableStateOf(false) }
    var isUnderlined by remember { mutableStateOf(false) }

    val context = LocalContext.current
    val imagePickerLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.GetContent()
    ) { uri: Uri? ->
        uri?.let { context.contentResolver.openInputStream(it)?.use { stream -> onImageSelected(BitmapFactory.decodeStream(stream)) } }
    }

    Surface(modifier = Modifier.fillMaxWidth(), color = MaterialTheme.colorScheme.surface, shadowElevation = 4.dp) {
        Row(
            modifier = Modifier.fillMaxWidth().padding(horizontal = 8.dp, vertical = 4.dp),
            verticalAlignment = Alignment.CenterVertically,
            horizontalArrangement = Arrangement.spacedBy(8.dp)
        ) {
            val isEnabled = currentSelection != null && !currentSelection.collapsed
            IconButton(
                onClick = { isBold = !isBold; onStyleChange(SpanStyle(fontWeight = FontWeight.Bold), isBold) },
                enabled = isEnabled
            ) {
                Image(painter = painterResource(id = R.drawable.bold), contentDescription = "Bold", alpha = if (isEnabled) 1f else 0.5f)
            }
            IconButton(
                onClick = { isItalic = !isItalic; onStyleChange(SpanStyle(fontStyle = FontStyle.Italic), isItalic) },
                enabled = isEnabled
            ) {
                Image(painter = painterResource(id = R.drawable.bold), contentDescription = "Italic", alpha = if (isEnabled) 1f else 0.5f)
            }
            IconButton(
                onClick = { isUnderlined = !isUnderlined; onStyleChange(SpanStyle(textDecoration = TextDecoration.Underline), isUnderlined) },
                enabled = isEnabled
            ) {
                Image(painter = painterResource(id = R.drawable.bold), contentDescription = "Underline", alpha = if (isEnabled) 1f else 0.5f)
            }
            Spacer(modifier = Modifier.weight(1f))
            IconButton(onClick = { imagePickerLauncher.launch("image/*") }) {
                Image(painter = painterResource(id = R.drawable.bold), contentDescription = "Insert Image")
            }
        }
    }
}