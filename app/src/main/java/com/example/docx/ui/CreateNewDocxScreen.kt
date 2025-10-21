package com.example.docx.ui

import android.content.Context
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
import androidx.compose.foundation.rememberScrollState
import androidx.compose.foundation.shape.RoundedCornerShape
import androidx.compose.foundation.text.BasicTextField
import androidx.compose.foundation.text.selection.TextSelectionColors
import androidx.compose.foundation.verticalScroll
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.draw.shadow
import androidx.compose.ui.focus.onFocusChanged
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.ColorFilter
import androidx.compose.ui.graphics.asImageBitmap
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.res.painterResource
import androidx.compose.ui.text.SpanStyle
import androidx.compose.ui.text.TextRange
import androidx.compose.ui.text.TextStyle
import androidx.compose.ui.text.buildAnnotatedString
import androidx.compose.ui.text.font.FontFamily
import androidx.compose.ui.text.font.FontStyle
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.input.TextFieldValue
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.text.style.TextDecoration
import androidx.compose.ui.text.withStyle
import androidx.compose.ui.unit.dp
import androidx.compose.ui.unit.sp
import androidx.navigation.NavController
import com.example.docx.R
import com.example.docx.util.Logger
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.util.Units
import org.apache.poi.xwpf.usermodel.*
import java.io.ByteArrayOutputStream
import java.io.IOException

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun CreateNewDocxScreen(navController: NavController) {
    var documentElements by remember {
        mutableStateOf<List<DocumentElement>>(
            listOf(DocumentElement.Paragraph(TextFieldValue("")))
        )
    }
    var focusedElementIndex by remember { mutableStateOf(-1) }
    var isSaving by remember { mutableStateOf(false) }
    var errorMessage by remember { mutableStateOf<String?>(null) }
    var showErrorDialog by remember { mutableStateOf(false) }
    var currentFont by remember { mutableStateOf("Times New Roman") }
    var currentFontSize by remember { mutableStateOf(12) }

    val context = LocalContext.current
    val scope = rememberCoroutineScope()

    val createDocumentLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
        onResult = { newFileUri: Uri? ->
            if (newFileUri != null) {
                scope.launch {
                    isSaving = true
                    val result = saveDocumentToFile(context, newFileUri, documentElements)
                    result.fold(
                        onSuccess = {
                            Logger.i("Document saved successfully to $newFileUri")
                            navController.popBackStack()
                        },
                        onFailure = { error ->
                            errorMessage = "Save failed: ${error.message}"
                            showErrorDialog = true
                        }
                    )
                    isSaving = false
                }
            } else {
                isSaving = false
                Logger.d("Save As action cancelled by user.")
            }
        }
    )

    val imagePickerLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.GetContent()
    ) { uri: Uri? ->
        uri?.let {
            context.contentResolver.openInputStream(it)?.use { stream ->
                val bitmap = BitmapFactory.decodeStream(stream)
                val newElements = documentElements.toMutableList()
                newElements.add(DocumentElement.Image(bitmap))
                documentElements = newElements
            }
        }
    }

    if (showErrorDialog) {
        AlertDialog(
            onDismissRequest = { 
                showErrorDialog = false 
                isSaving = false
            },
            title = { Text("Error") },
            text = { Text(errorMessage ?: "An unknown error occurred.") },
            confirmButton = {
                TextButton(onClick = { 
                    showErrorDialog = false 
                    isSaving = false
                }) { Text("OK") }
            }
        )
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = { Text("Create New Document") },
                navigationIcon = {
                    IconButton(onClick = { navController.popBackStack() }, enabled = !isSaving) {
                        Icon(Icons.AutoMirrored.Filled.ArrowBack, contentDescription = "Back")
                    }
                },
                actions = {
                    TextButton(
                        onClick = {
                            isSaving = true
                            createDocumentLauncher.launch("Untitled.docx")
                        },
                        enabled = !isSaving
                    ) {
                        Text(if (isSaving) "Saving..." else "Save")
                    }
                }
            )
        }
    ) { paddingValues ->
        Column(
            modifier = Modifier
                .fillMaxSize()
                .padding(paddingValues)
        ) {
            // Formatting toolbar
            FormattingToolbar(
                currentFont = currentFont,
                onFontChange = { font ->
                    currentFont = font
                    applyFormatting(
                        documentElements, focusedElementIndex,
                        SpanStyle(fontFamily = getFontFamily(font))
                    ) { newElements -> documentElements = newElements }
                },
                currentFontSize = currentFontSize,
                onFontSizeChange = { size ->
                    currentFontSize = size
                    applyFormatting(
                        documentElements, focusedElementIndex,
                        SpanStyle(fontSize = size.sp)
                    ) { newElements -> documentElements = newElements }
                },
                onStyleChange = { style ->
                    applyFormatting(
                        documentElements, focusedElementIndex, style
                    ) { newElements -> documentElements = newElements }
                },
                onInsertImage = { imagePickerLauncher.launch("image/*") },
                onInsertTable = {
                    val newElements = documentElements.toMutableList()
                    val table = DocumentElement.Table()
                    // Initialize with 2x2 table
                    repeat(2) {
                        val row = mutableListOf<TextFieldValue>()
                        repeat(2) { row.add(TextFieldValue("")) }
                        table.rows.add(row)
                    }
                    newElements.add(table)
                    documentElements = newElements
                },
                currentSelection = getCurrentSelection(documentElements, focusedElementIndex)
            )

            // Document area with Word-like appearance
            Box(
                modifier = Modifier
                    .weight(1f)
                    .background(Color(0xFFF5F5F5))
                    .padding(32.dp)
            ) {
                // Word document page
                Card(
                    modifier = Modifier
                        .fillMaxWidth()
                        .defaultMinSize(minHeight = 800.dp)
                        .shadow(8.dp, RoundedCornerShape(4.dp)),
                    colors = CardDefaults.cardColors(containerColor = Color.White),
                    shape = RoundedCornerShape(4.dp)
                ) {
                    LazyColumn(
                        modifier = Modifier
                            .padding(
                                top = 72.dp,    // 1 inch top margin
                                bottom = 72.dp, // 1 inch bottom margin
                                start = 72.dp,  // 1 inch left margin
                                end = 72.dp     // 1 inch right margin
                            )
                            .fillMaxWidth(),
                        verticalArrangement = Arrangement.spacedBy(8.dp)
                    ) {
                        itemsIndexed(documentElements) { index, element ->
                            when (element) {
                                is DocumentElement.Paragraph -> {
                                    ParagraphEditor(
                                        paragraph = element,
                                        onValueChange = { newValue ->
                                            val newElements = documentElements.toMutableList()
                                            (newElements[index] as DocumentElement.Paragraph).content =
                                                newValue
                                            documentElements = newElements
                                        },
                                        onFocusChanged = { focused ->
                                            if (focused) focusedElementIndex = index
                                        },
                                        modifier = Modifier.fillMaxWidth()
                                    )
                                }
                                is DocumentElement.Image -> {
                                    ImageElement(
                                        image = element,
                                        modifier = Modifier.fillMaxWidth()
                                    )
                                }

                                is DocumentElement.Table -> {
                                    TableEditor(
                                        table = element,
                                        onCellValueChange = { rowIndex, colIndex, newValue ->
                                            val newElements = documentElements.toMutableList()
                                            val table = newElements[index] as DocumentElement.Table
                                            table.rows[rowIndex][colIndex] = newValue
                                            documentElements = newElements
                                        },
                                        modifier = Modifier.fillMaxWidth()
                                    )
                                }
                            }
                        }
                    }
                }

                if (isSaving) {
                    Surface(
                        modifier = Modifier.fillMaxSize(),
                        color = MaterialTheme.colorScheme.surface.copy(alpha = 0.8f)
                    ) {
                        Column(
                            modifier = Modifier.fillMaxSize(),
                            horizontalAlignment = Alignment.CenterHorizontally,
                            verticalArrangement = Arrangement.Center
                        ) {
                            CircularProgressIndicator()
                            Spacer(modifier = Modifier.height(16.dp))
                            Text("Preparing to save...")
                        }
                    }
                }
            }
        }
    }
}

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun FormattingToolbar(
    currentFont: String,
    onFontChange: (String) -> Unit,
    currentFontSize: Int,
    onFontSizeChange: (Int) -> Unit,
    onStyleChange: (SpanStyle) -> Unit,
    onInsertImage: () -> Unit,
    onInsertTable: () -> Unit,
    currentSelection: TextRange?
) {
    val fonts = listOf("Arial", "Calibri", "Courier New", "Georgia", "Times New Roman", "Verdana")
    val fontSizes = listOf(8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72)

    var isFontDropdownExpanded by remember { mutableStateOf(false) }
    var isFontSizeDropdownExpanded by remember { mutableStateOf(false) }
    var isBoldSelected by remember { mutableStateOf(false) }
    var isItalicSelected by remember { mutableStateOf(false) }
    var isUnderlineSelected by remember { mutableStateOf(false) }

    Column {
        // Font and size row
        Row(
            modifier = Modifier
                .fillMaxWidth()
                .padding(horizontal = 8.dp, vertical = 4.dp),
            verticalAlignment = Alignment.CenterVertically,
            horizontalArrangement = Arrangement.spacedBy(8.dp)
        ) {
            // Font dropdown
            ExposedDropdownMenuBox(
                expanded = isFontDropdownExpanded,
                onExpandedChange = { isFontDropdownExpanded = !isFontDropdownExpanded },
                modifier = Modifier.weight(1f)
            ) {
                OutlinedTextField(
                    value = currentFont,
                    onValueChange = {},
                    readOnly = true,
                    label = { Text("Font") },
                    trailingIcon = { ExposedDropdownMenuDefaults.TrailingIcon(expanded = isFontDropdownExpanded) },
                    modifier = Modifier.menuAnchor()
                )
                ExposedDropdownMenu(
                    expanded = isFontDropdownExpanded,
                    onDismissRequest = { isFontDropdownExpanded = false }
                ) {
                    fonts.forEach { fontName ->
                        DropdownMenuItem(
                            text = { Text(fontName) },
                            onClick = {
                                onFontChange(fontName)
                                isFontDropdownExpanded = false
                            }
                        )
                    }
                }
            }

            // Font size dropdown
            ExposedDropdownMenuBox(
                expanded = isFontSizeDropdownExpanded,
                onExpandedChange = { isFontSizeDropdownExpanded = !isFontSizeDropdownExpanded },
                modifier = Modifier.width(80.dp)
            ) {
                OutlinedTextField(
                    value = currentFontSize.toString(),
                    onValueChange = {},
                    readOnly = true,
                    label = { Text("Size") },
                    trailingIcon = { ExposedDropdownMenuDefaults.TrailingIcon(expanded = isFontSizeDropdownExpanded) },
                    modifier = Modifier.menuAnchor()
                )
                ExposedDropdownMenu(
                    expanded = isFontSizeDropdownExpanded,
                    onDismissRequest = { isFontSizeDropdownExpanded = false }
                ) {
                    fontSizes.forEach { size ->
                        DropdownMenuItem(
                            text = { Text(size.toString()) },
                            onClick = {
                                onFontSizeChange(size)
                                isFontSizeDropdownExpanded = false
                            }
                        )
                    }
                }
            }
        }

        // Style and tools row
        Row(
            modifier = Modifier
                .fillMaxWidth()
                .padding(horizontal = 8.dp, vertical = 4.dp),
            verticalAlignment = Alignment.CenterVertically,
            horizontalArrangement = Arrangement.spacedBy(8.dp)
        ) {
            val isEnabled = currentSelection != null && !currentSelection.collapsed

            IconButton(
                onClick = {
                    isBoldSelected = !isBoldSelected
                    onStyleChange(SpanStyle(fontWeight = FontWeight.Bold))
                },
                enabled = isEnabled
            ) {
                Image(
                    painter = painterResource(id = R.drawable.bold),
                    contentDescription = "Bold",
                    colorFilter = if (isBoldSelected) ColorFilter.tint(MaterialTheme.colorScheme.primary) else null
                )
            }

            IconButton(
                onClick = {
                    isItalicSelected = !isItalicSelected
                    onStyleChange(SpanStyle(fontStyle = FontStyle.Italic))
                },
                enabled = isEnabled
            ) {
                Icon(
                    painter = painterResource(R.drawable.bold),
                    contentDescription = "Italic",
                    tint = if (isItalicSelected) MaterialTheme.colorScheme.primary else Color.Unspecified
                )
            }

            IconButton(
                onClick = {
                    isUnderlineSelected = !isUnderlineSelected
                    onStyleChange(SpanStyle(textDecoration = TextDecoration.Underline))
                },
                enabled = isEnabled
            ) {
                Icon(
                    painter = painterResource(R.drawable.bold),
                    contentDescription = "Underline",
                    tint = if (isUnderlineSelected) MaterialTheme.colorScheme.primary else Color.Unspecified
                )
            }

            Spacer(modifier = Modifier.weight(1f))

            IconButton(onClick = onInsertImage) {
                Icon(
                    painter = painterResource(R.drawable.bold),
                    contentDescription = "Insert Image"
                )
            }

            IconButton(onClick = onInsertTable) {
                Icon(
                    painter = painterResource(R.drawable.bold),
                    contentDescription = "Insert Table"
                )
            }
        }
    }
}

// Document element types
sealed class DocumentElement {
    data class Paragraph(
        var content: TextFieldValue,
        var alignment: ParagraphAlignment = ParagraphAlignment.LEFT,
        var spacingBefore: Int = 0,
        var spacingAfter: Int = 0,
        var lineSpacing: Double = 1.0
    ) : DocumentElement()

    data class Image(
        val bitmap: Bitmap,
        val width: Int = 200,
        val height: Int = 150,
        var alignment: ParagraphAlignment = ParagraphAlignment.CENTER
    ) : DocumentElement()

    data class Table(
        val rows: MutableList<MutableList<TextFieldValue>> = mutableListOf(),
        val columnCount: Int = 2
    ) : DocumentElement()
}

// --- Compose UI Editors ---

@Composable
fun ParagraphEditor(
    paragraph: DocumentElement.Paragraph,
    onValueChange: (TextFieldValue) -> Unit,
    onFocusChanged: (Boolean) -> Unit,
    modifier: Modifier = Modifier
) {
    BasicTextField(
        value = paragraph.content,
        onValueChange = onValueChange,
        textStyle = TextStyle(
            fontSize = 12.sp,
            fontFamily = FontFamily.Serif,
            lineHeight = (12 * paragraph.lineSpacing).sp,
            textAlign = when (paragraph.alignment) {
                ParagraphAlignment.LEFT -> TextAlign.Left
                ParagraphAlignment.CENTER -> TextAlign.Center
                ParagraphAlignment.RIGHT -> TextAlign.Right
                ParagraphAlignment.BOTH -> TextAlign.Justify
                else -> TextAlign.Left
            }
        ),
        modifier = modifier
            .fillMaxWidth()
            .padding(vertical = 4.dp)
            .onFocusChanged { onFocusChanged(it.isFocused) },
        decorationBox = { innerTextField ->
            Box(
                modifier = Modifier
                    .fillMaxWidth()
                    .padding(4.dp)
            ) {
                if (paragraph.content.text.isEmpty()) {
                    Text(
                        text = "Type here...",
                        style = TextStyle(
                            fontSize = 12.sp,
                            fontFamily = FontFamily.Serif,
                            color = Color.Gray
                        )
                    )
                }
                innerTextField()
            }
        }
    )
}

@Composable
fun ImageElement(
    image: DocumentElement.Image,
    modifier: Modifier = Modifier
) {
    Box(
        modifier = modifier,
        contentAlignment = when (image.alignment) {
            ParagraphAlignment.CENTER -> Alignment.Center
            ParagraphAlignment.RIGHT -> Alignment.CenterEnd
            else -> Alignment.CenterStart
        }
    ) {
        Image(
            bitmap = image.bitmap.asImageBitmap(),
            contentDescription = "Document Image",
            modifier = Modifier.size(image.width.dp, image.height.dp)
        )
    }
}

@Composable
fun TableEditor(
    table: DocumentElement.Table,
    onCellValueChange: (Int, Int, TextFieldValue) -> Unit,
    modifier: Modifier = Modifier
) {
    Column(modifier = modifier) {
        table.rows.forEachIndexed { rowIndex, row ->
            Row(
                modifier = Modifier.fillMaxWidth(),
                horizontalArrangement = Arrangement.spacedBy(1.dp)
            ) {
                row.forEachIndexed { colIndex, cellValue ->
                    BasicTextField(
                        value = cellValue,
                        onValueChange = { onCellValueChange(rowIndex, colIndex, it) },
                        textStyle = TextStyle(fontSize = 12.sp, fontFamily = FontFamily.Serif),
                        modifier = Modifier
                            .weight(1f)
                            .border(1.dp, Color.Black)
                            .padding(8.dp),
                        decorationBox = { innerTextField ->
                            Box(modifier = Modifier.fillMaxWidth()) {
                                if (cellValue.text.isEmpty()) {
                                    Text(
                                        text = "Cell",
                                        style = TextStyle(fontSize = 12.sp, color = Color.Gray)
                                    )
                                }
                                innerTextField()
                            }
                        }
                    )
                }
            }
        }
    }
}

// Helper functions
private fun getFontFamily(fontName: String): FontFamily {
    return when (fontName) {
        "Arial", "Verdana", "Calibri" -> FontFamily.SansSerif
        "Courier New" -> FontFamily.Monospace
        else -> FontFamily.Serif
    }
}

private fun getCurrentSelection(elements: List<DocumentElement>, focusedIndex: Int): TextRange? {
    if (focusedIndex == -1 || focusedIndex >= elements.size) return null
    val element = elements[focusedIndex]
    return if (element is DocumentElement.Paragraph) element.content.selection else null
}

private fun applyFormatting(
    elements: List<DocumentElement>,
    focusedIndex: Int,
    style: SpanStyle,
    onUpdate: (List<DocumentElement>) -> Unit
) {
    if (focusedIndex == -1 || focusedIndex >= elements.size) return

    val element = elements[focusedIndex]
    if (element is DocumentElement.Paragraph) {
        val selection = element.content.selection
        if (!selection.collapsed) {
            val newElements = elements.toMutableList()
            val paragraph = newElements[focusedIndex] as DocumentElement.Paragraph
            val annotatedString = buildAnnotatedString {
                append(paragraph.content.annotatedString)
                addStyle(style, selection.start, selection.end)
            }
            paragraph.content = TextFieldValue(annotatedString, selection)
            onUpdate(newElements)
        }
    }
}

// Save document using Apache POI
suspend fun saveDocumentToFile(
    context: Context,
    uri: Uri,
    elements: List<DocumentElement>
): Result<Unit> {
    return withContext(Dispatchers.IO) {
        runCatching {
            context.contentResolver.openOutputStream(uri)?.use { outputStream ->
                val document = XWPFDocument()

                // Set A4 page size and margins (same as original)
                setA4PageSize(document)

                elements.forEach { element ->
                    when (element) {
                        is DocumentElement.Paragraph -> {
                            val paragraph = document.createParagraph()
                            paragraph.alignment = element.alignment
                            paragraph.spacingBefore = element.spacingBefore
                            paragraph.spacingAfter = element.spacingAfter

                            val annotatedString = element.content.annotatedString
                            val text = annotatedString.text

                            if (text.isNotEmpty()) {
                                // Handle styled text similar to EditDocxScreen
                                val boundaries = mutableSetOf(0, text.length)
                                annotatedString.spanStyles.forEach {
                                    boundaries.add(it.start)
                                    boundaries.add(it.end)
                                }

                                val sortedBoundaries =
                                    boundaries.filter { it < text.length }.sorted()

                                for (i in 0 until sortedBoundaries.size) {
                                    val start = sortedBoundaries[i]
                                    val end =
                                        if (i < sortedBoundaries.size - 1) sortedBoundaries[i + 1] else text.length
                                    if (start >= end) continue

                                    val run = paragraph.createRun()
                                    run.setText(text.substring(start, end))

                                    val styles =
                                        annotatedString.spanStyles.filter { spanStyleRange ->
                                            spanStyleRange.start <= start && spanStyleRange.end > start
                                        }

                                    styles.forEach { spanStyleRange ->
                                        val style = spanStyleRange.item
                                        if (style.fontWeight == FontWeight.Bold) run.isBold = true
                                        if (style.fontStyle == FontStyle.Italic) run.isItalic = true
                                        if (style.textDecoration == TextDecoration.Underline) run.setUnderline(
                                            UnderlinePatterns.SINGLE
                                        )
                                        style.fontSize?.let { fontSize ->
                                            run.fontSize = fontSize.value.toInt()
                                        }
                                    }
                                }
                            } else {
                                paragraph.createRun() // Empty paragraph
                            }
                        }

                        is DocumentElement.Image -> {
                            val paragraph = document.createParagraph()
                            paragraph.alignment = element.alignment
                            val run = paragraph.createRun()
                            val stream = ByteArrayOutputStream()
                            element.bitmap.compress(Bitmap.CompressFormat.PNG, 100, stream)
                            val pictureType = XWPFDocument.PICTURE_TYPE_PNG
                            run.addPicture(
                                stream.toByteArray().inputStream(),
                                pictureType,
                                "image.png",
                                Units.toEMU(element.width.toDouble()),
                                Units.toEMU(element.height.toDouble())
                            )
                        }

                        is DocumentElement.Table -> {
                            val table = document.createTable()
                            element.rows.forEachIndexed { rowIndex, row ->
                                val tableRow =
                                    if (rowIndex == 0) table.getRow(0) else table.createRow()
                                row.forEachIndexed { colIndex, cellValue ->
                                    val cell = if (colIndex < tableRow.tableCells.size) {
                                        tableRow.getCell(colIndex)
                                    } else {
                                        tableRow.createCell()
                                    }
                                    cell.text = cellValue.text
                                }
                            }
                        }
                    }
                }

                document.write(outputStream)
                document.close()
                Logger.i("DOCX saved successfully to $uri")
            } ?: throw IOException("Unable to open output stream for $uri")
        }.onFailure { e ->
            Logger.e("Error saving document: ${e.message}", e)
        }
    }
}

// Set A4 page size and margins (same as original)
private fun setA4PageSize(document: XWPFDocument) {
    try {
        val ctDocument = document.document
        val ctBody = ctDocument.body

        val sectPr = if (ctBody.isSetSectPr) ctBody.sectPr else ctBody.addNewSectPr()

        // Set page size to A4 (210mm x 297mm = 11906 twips x 16838 twips)
        val pgSz = if (sectPr.isSetPgSz) sectPr.pgSz else sectPr.addNewPgSz()
        pgSz.w = java.math.BigInteger.valueOf(11906) // A4 width in twips
        pgSz.h = java.math.BigInteger.valueOf(16838) // A4 height in twips

        // Set margins (2.54cm = 1440 twips)
        val pgMar = if (sectPr.isSetPgMar) sectPr.pgMar else sectPr.addNewPgMar()
        pgMar.top = java.math.BigInteger.valueOf(1440)    // 2.54cm
        pgMar.bottom = java.math.BigInteger.valueOf(1440) // 2.54cm
        pgMar.left = java.math.BigInteger.valueOf(1440)   // 2.54cm
        pgMar.right = java.math.BigInteger.valueOf(1440)  // 2.54cm

    } catch (e: Exception) {
        Logger.w("Failed to set A4 page size: ${e.message}")
    }
}