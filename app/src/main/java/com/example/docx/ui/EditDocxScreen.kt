package com.example.docx.ui

import android.annotation.SuppressLint
import android.content.Context
import android.net.Uri
import android.util.Base64
import android.webkit.WebView
import androidx.activity.compose.rememberLauncherForActivityResult
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.Image
import androidx.compose.foundation.layout.*
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.graphics.ColorFilter
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.res.painterResource
import androidx.compose.ui.unit.dp
import androidx.compose.ui.viewinterop.AndroidView
import androidx.documentfile.provider.DocumentFile
import androidx.navigation.NavController
import com.example.docx.R
import com.example.docx.util.Logger
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.common.usermodel.PictureType
import org.apache.poi.util.Units
import org.apache.poi.xwpf.usermodel.*
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STVerticalAlignRun as OfficeSTVerticalAlignRun
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*
import org.jsoup.Jsoup
import org.jsoup.nodes.Element
import org.jsoup.nodes.Node
import org.jsoup.nodes.TextNode
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.io.IOException
import java.math.BigInteger

// =================================================================================
// EditDocxScreen Composable
// =================================================================================

@OptIn(ExperimentalMaterial3Api::class)
@SuppressLint("SetJavaScriptEnabled")
@Composable
fun EditDocxScreen(navController: NavController, fileUriString: String) {
    var documentHtml by remember { mutableStateOf("<p></p>") }
    var isLoading by remember { mutableStateOf(true) }
    var isSaving by remember { mutableStateOf(false) }
    val context = LocalContext.current
    val scope = rememberCoroutineScope()
    var errorMessage by remember { mutableStateOf<String?>(null) }
    var showErrorDialog by remember { mutableStateOf(false) }
    val webView = remember { WebView(context) }

    val currentFileUri = remember { Uri.parse(fileUriString) }

    val saveAsLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
        onResult = { newFileUri: Uri? ->
            if (newFileUri != null) {
                webView.evaluateJavascript(
                    "(function() { return document.documentElement.outerHTML; })();"
                ) { htmlContent ->
                    scope.launch {
                        isSaving = true
                        val result = writeHtmlToDocx(context, newFileUri, htmlContent ?: "<html><body></body></html>")
                        result.fold(
                            onSuccess = {
                                Logger.i("Document saved successfully to $newFileUri")
                                navController.navigateUp()
                            },
                            onFailure = { error ->
                                errorMessage = "Save failed: ${error.message}"
                                showErrorDialog = true
                            }
                        )
                        if(!showErrorDialog) isSaving = false
                    }
                }
            } else {
                isSaving = false
                Logger.d("Save As action cancelled.")
            }
        }
    )

    LaunchedEffect(currentFileUri) {
        isLoading = true
        Logger.d("LaunchedEffect: Loading DOCX from $currentFileUri")
        withContext(Dispatchers.IO) {
            try {
                context.contentResolver.openInputStream(currentFileUri)?.use { inputStream ->
                    val document = XWPFDocument(inputStream)
                    val html = convertDocxToHtmlEnhanced(document)
                    withContext(Dispatchers.Main) {
                        documentHtml = html
                        Logger.d("Document loaded and converted to HTML successfully.")
                    }
                    document.close()
                } ?: throw IOException("Failed to open input stream for $currentFileUri")
            } catch (e: Exception) {
                Logger.e("Error loading DOCX: ${e.localizedMessage}", e)
                withContext(Dispatchers.Main) {
                    errorMessage = "Load failed: ${e.localizedMessage}"
                    showErrorDialog = true
                }
            } finally {
                withContext(Dispatchers.Main) {
                    isLoading = false
                }
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
                    if (errorMessage?.startsWith("Load failed") == true) navController.navigateUp()
                }) { Text("OK") }
            }
        )
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = { Text(DocumentFile.fromSingleUri(context, currentFileUri)?.name ?: "Edit Document") },
                navigationIcon = {
                    IconButton(onClick = { navController.navigateUp() }, enabled = !isSaving) {
                        Icon(Icons.AutoMirrored.Filled.ArrowBack, "Back")
                    }
                },
                actions = {
                    TextButton(
                        onClick = {
                            isSaving = true
                            val originalFileName = DocumentFile.fromSingleUri(context, currentFileUri)?.name ?: "Untitled.docx"
                            saveAsLauncher.launch(originalFileName)
                        },
                        enabled = !isLoading && !isSaving
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
            FormattingToolbar(webView = webView)
            Box(modifier = Modifier.weight(1f)) {
                if (isLoading) {
                    CircularProgressIndicator(modifier = Modifier.align(Alignment.Center))
                } else {
                    AndroidView(
                        factory = { webViewContext ->
                            webView.apply {
                                settings.javaScriptEnabled = true
                                settings.domStorageEnabled = true
                                settings.allowFileAccess = true
                                settings.defaultTextEncodingName = "utf-8"
                                settings.builtInZoomControls = true
                                settings.displayZoomControls = false

                                webViewClient = object : android.webkit.WebViewClient() {
                                    override fun onPageFinished(view: WebView?, url: String?) {
                                        evaluateJavascript("document.body.contentEditable = true;") {}
                                    }
                                }
                                loadDataWithBaseURL(null, documentHtml, "text/html", "utf-8", null)
                            }
                        },
                        modifier = Modifier
                            .fillMaxSize()
                            .padding(horizontal = 16.dp)
                    )
                }
                if (isSaving) {
                    Surface(modifier = Modifier.fillMaxSize(), color = MaterialTheme.colorScheme.surface.copy(alpha = 0.8f)) {
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
fun FormattingToolbar(webView: WebView) {
    val fonts = listOf("Arial", "Calibri", "Courier New", "Georgia", "Times New Roman", "Verdana")
    var selectedFont by remember { mutableStateOf(fonts[0]) }
    var isFontDropdownExpanded by remember { mutableStateOf(false) }

    var isBoldSelected by remember { mutableStateOf(false) }
    var isItalicSelected by remember { mutableStateOf(false) }
    var isUnderlineSelected by remember { mutableStateOf(false) }

    Row(
        modifier = Modifier
            .fillMaxWidth()
            .padding(horizontal = 8.dp, vertical = 4.dp),
        verticalAlignment = Alignment.CenterVertically,
        horizontalArrangement = Arrangement.spacedBy(8.dp)
    ) {
        ExposedDropdownMenuBox(
            expanded = isFontDropdownExpanded,
            onExpandedChange = { isFontDropdownExpanded = !isFontDropdownExpanded },
            modifier = Modifier.weight(1f)
        ) {
            OutlinedTextField(
                value = selectedFont,
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
                            selectedFont = fontName
                            isFontDropdownExpanded = false
                            webView.evaluateJavascript("document.execCommand('fontName', false, '$fontName');", null)
                        }
                    )
                }
            }
        }

        IconButton(onClick = {
            isBoldSelected = !isBoldSelected
            webView.evaluateJavascript("document.execCommand('bold');", null)
        }) {
            Image(
                painter = painterResource(id = R.drawable.bold),
                contentDescription = "Bold",
                colorFilter = if (isBoldSelected) ColorFilter.tint(MaterialTheme.colorScheme.primary) else null
            )
        }

        IconButton(onClick = {
            isItalicSelected = !isItalicSelected
            webView.evaluateJavascript("document.execCommand('italic');", null)
        }) {
            Icon(
                painter = painterResource(R.drawable.bold),
                contentDescription = "Italic",
                tint = if (isItalicSelected) MaterialTheme.colorScheme.primary else Color.Unspecified
            )
        }

        IconButton(onClick = {
            isUnderlineSelected = !isUnderlineSelected
            webView.evaluateJavascript("document.execCommand('underline');", null)
        }) {
            Icon(
                painter = painterResource(R.drawable.bold),
                contentDescription = "Underline",
                tint = if (isUnderlineSelected) MaterialTheme.colorScheme.primary else Color.Unspecified
            )
        }
    }
}

// =================================================================================
// DOCX -> HTML CONVERSION (Enhanced)
// =================================================================================

fun convertDocxToHtmlEnhanced(document: XWPFDocument): String {
    val html = StringBuilder()
    html.append("<!DOCTYPE html><html><head>")
    html.append("<meta charset=\"UTF-8\">")
    html.append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0, user-scalable=yes\">")
    html.append("<style>")
    html.append("body { ")
    html.append("    font-family: 'Times New Roman', serif; ")
    html.append("    line-height: 1.6; ")
    html.append("    background: #f0f0f0; ")
    html.append("    margin: 0; padding: 20px; }")
    html.append(".page { ")
    html.append("    background: white; ")
    html.append("    margin: 0 auto 20px auto; ")
    html.append("    padding-top: 1.00cm; ")
    html.append("    padding-bottom: 1.00cm; ")
    html.append("    padding-left: 2.54cm; ")
    html.append("    padding-right: 2.54cm; ")
    html.append("    width: 21cm; ")
    html.append("    min-height: 29.7cm; ")
    html.append("    box-shadow: 0 4px 8px rgba(0,0,0,0.15); ")
    html.append("    page-break-after: always; ")
    html.append("    position: relative; ")
    html.append("    overflow: visible; }")
    html.append(".page:last-child { page-break-after: auto; }")
    html.append(".page-number { position: absolute; bottom: 1.27cm; right: 2.54cm; font-size: 10pt; color: #666; }")
    html.append("h1, h2, h3, h4, h5 { margin-top: 0.4em; margin-bottom: 0.2em; }")
    html.append("p { margin: 0.2em 0; }")
    html.append("ul, ol { margin: 0.5em 0 0.5em 1em; }")
    html.append("table { border-collapse: collapse; width: 100%; margin: 1em 0; }")
    html.append("td, th { border: 1px solid #444; padding: 3px; text-align: left; }")
    html.append("img { max-width: 100%; height: auto; display: block; margin: 0.5em auto; }")
    html.append("</style>")
    html.append("<script>")
    html.append("let pageCounter = 1;")
    html.append("let isProcessing = false;")
    html.append("")
    html.append("function createNewPage() {")
    html.append("  pageCounter++;")
    html.append("  const newPage = document.createElement('div');")
    html.append("  newPage.className = 'page';")
    html.append("  newPage.id = 'page' + pageCounter;")
    html.append("  newPage.innerHTML = '<p><br></p>';")
    html.append("  ")
    html.append("  // Add page number")
    html.append("  const pageNum = document.createElement('div');")
    html.append("  pageNum.className = 'page-number';")
    html.append("  pageNum.textContent = pageCounter;")
    html.append("  newPage.appendChild(pageNum);")
    html.append("  ")
    html.append("  document.body.appendChild(newPage);")
    html.append("  return newPage;")
    html.append("}")
    html.append("")
    html.append("function updatePageNumbers() {")
    html.append("  const pages = document.querySelectorAll('.page');")
    html.append("  pages.forEach((page, index) => {")
    html.append("    let pageNum = page.querySelector('.page-number');")
    html.append("    if (!pageNum) {")
    html.append("      pageNum = document.createElement('div');")
    html.append("      pageNum.className = 'page-number';")
    html.append("      page.appendChild(pageNum);")
    html.append("    }")
    html.append("    pageNum.textContent = index + 1;")
    html.append("  });")
    html.append("}")
    html.append("")
    html.append("function checkPageOverflow() {")
    html.append("  if (isProcessing) return;")
    html.append("  isProcessing = true;")
    html.append("  ")
    html.append("  const pages = document.querySelectorAll('.page');")
    html.append("  const pageHeight = 29.7 * 37.795; // A4 height in pixels")
    html.append("  ")
    html.append("  pages.forEach((page, index) => {")
    html.append("    const contentHeight = page.scrollHeight;")
    html.append("    if (contentHeight > pageHeight) {")
    html.append("      // Create new page if it doesn't exist")
    html.append("      let nextPage = document.getElementById('page' + (index + 2));")
    html.append("      if (!nextPage) {")
    html.append("        nextPage = createNewPage();")
    html.append("      }")
    html.append("      ")
    html.append("      // Move overflow content to next page")
    html.append("      const elements = Array.from(page.children);")
    html.append("      let totalHeight = 0;")
    html.append("      let moveIndex = -1;")
    html.append("      ")
    html.append("      for (let i = 0; i < elements.length; i++) {")
    html.append("        const elem = elements[i];")
    html.append("        if (elem.className === 'page-number') continue;")
    html.append("        ")
    html.append("        const elemHeight = elem.offsetHeight;")
    html.append("        if (totalHeight + elemHeight > pageHeight - 100) {") // Leave some margin
    html.append("          moveIndex = i;")
    html.append("          break;")
    html.append("        }")
    html.append("        totalHeight += elemHeight;")
    html.append("      }")
    html.append("      ")
    html.append("      if (moveIndex !== -1) {")
    html.append("        for (let i = moveIndex; i < elements.length; i++) {")
    html.append("          const elem = elements[i];")
    html.append("          if (elem.className !== 'page-number') {")
    html.append("            nextPage.insertBefore(elem, nextPage.querySelector('.page-number'));")
    html.append("          }")
    html.append("        }")
    html.append("      }")
    html.append("    }")
    html.append("  });")
    html.append("  ")
    html.append("  updatePageNumbers();")
    html.append("  isProcessing = false;")
    html.append("}")
    html.append("")
    html.append("function ensureMinimumOnePage() {")
    html.append("  if (document.querySelectorAll('.page').length === 0) {")
    html.append("    const firstPage = document.createElement('div');")
    html.append("    firstPage.className = 'page';")
    html.append("    firstPage.id = 'page1';")
    html.append("    firstPage.innerHTML = '<p><br></p>';")
    html.append("    document.body.appendChild(firstPage);")
    html.append("  }")
    html.append("  updatePageNumbers();")
    html.append("}")
    html.append("")
    html.append("document.addEventListener('DOMContentLoaded', function() {")
    html.append("  ensureMinimumOnePage();")
    html.append("  setTimeout(checkPageOverflow, 500);")
    html.append("});")
    html.append("")
    html.append("document.addEventListener('input', function() {")
    html.append("  setTimeout(checkPageOverflow, 300);")
    html.append("});")
    html.append("")
    html.append("document.addEventListener('keydown', function(e) {")
    html.append("  if (e.ctrlKey && e.key === 'Enter') {")
    html.append("    e.preventDefault();")
    html.append("    createNewPage();")
    html.append("    updatePageNumbers();")
    html.append("  }")
    html.append("});")
    html.append("")
    html.append("// Periodic check for content changes")
    html.append("setInterval(checkPageOverflow, 2000);")
    html.append("</script>")
    html.append("</head><body>")

    html.append("<div class='page' id='page1'>")

    for (bodyElement in document.bodyElements) {
        when (bodyElement) {
            is XWPFParagraph -> html.append(convertParagraphToHtmlExact(bodyElement))
            is XWPFTable -> html.append(convertTableToHtmlExact(bodyElement))
        }
    }

    html.append("</div>") // Close first page
    html.append("</body></html>")
    return html.toString()
}

fun convertParagraphToHtmlExact(paragraph: XWPFParagraph): String {
    val style = buildString {
        if (paragraph.spacingBefore > 0) append("margin-top: ${paragraph.spacingBefore / 20.0}pt; ")
        if (paragraph.spacingAfter > 0) append("margin-bottom: ${paragraph.spacingAfter / 20.0}pt; ")
        if (paragraph.spacingLineRule == LineSpacingRule.AUTO && paragraph.spacingBetween > 0) {
            append("line-height: ${paragraph.spacingBetween / 240.0}; ")
        }
        if (paragraph.alignment != ParagraphAlignment.LEFT) {
            append("text-align: ${paragraph.alignment.name.lowercase()}; ")
        }
    }
    val pTag = StringBuilder()
    pTag.append("<p style='${style}'>")
    var contentAdded = false
    for (run in paragraph.runs) {
        val runHtml = convertRunToHtmlExact(run)
        if (runHtml.isNotBlank()) {
            pTag.append(runHtml)
            contentAdded = true
        }
    }
    if (!contentAdded && paragraph.text.isBlank()) {
        pTag.append("&nbsp;")
    }
    pTag.append("</p>")
    return pTag.toString()
}

fun convertRunToHtmlExact(run: XWPFRun): String {
    val style = buildString {
        val fontSizeInPoints: Double? = run.getFontSizeAsDouble()
        if (fontSizeInPoints != null && fontSizeInPoints > 0) {
            append("font-size: ${fontSizeInPoints}pt; ")
        }
        if (!run.fontFamily.isNullOrEmpty()) append("font-family: '${run.fontFamily}'; ")
        if (run.color != null && run.color != "auto") append("color: #${run.color}; ")
        if (run.isBold) append("font-weight: bold; ")
        if (run.isItalic) append("font-style: italic; ")
        if (run.isStrikeThrough) append("text-decoration: line-through; ")
        if (run.underline != UnderlinePatterns.NONE && run.underline != null) append("text-decoration: underline; ")
    }
    var text = run.text() ?: ""
    text = text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\"", "&quot;")
        .replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;")
        .replace("\n", "<br>")
    var html = "<span style='${style}'>$text</span>"
    when (run.getVerticalAlignment()) {
        OfficeSTVerticalAlignRun.SUBSCRIPT -> html = "<sub>$html</sub>"
        OfficeSTVerticalAlignRun.SUPERSCRIPT -> html = "<sup>$html</sup>"
        else -> {}
    }
    for (pic in run.embeddedPictures) {
        try {
            val picData = pic.pictureData
            val ext = when (picData.pictureTypeEnum) {
                PictureType.JPEG -> "jpeg"
                PictureType.PNG -> "png"
                PictureType.GIF -> "gif"
                PictureType.BMP -> "bmp"
                PictureType.WMF -> "wmf"
                PictureType.EMF -> "emf"
                PictureType.PICT -> "pict"
                PictureType.TIFF -> "tiff"
                else -> "png"
            }
            val base64 = Base64.encodeToString(picData.data, Base64.DEFAULT)
            val xfrm = pic.ctPicture?.spPr?.xfrm
            var imgWidth = "auto"
            var imgHeight = "auto"
            if (xfrm?.ext != null) {
                val widthEmu = xfrm.ext.cx
                val heightEmu = xfrm.ext.cy
                imgWidth = "${widthEmu / Units.EMU_PER_PIXEL}px"
                imgHeight = "${heightEmu / Units.EMU_PER_PIXEL}px"
            }
            html += "<img src='data:image/$ext;base64,$base64' width='$imgWidth' height='$imgHeight' alt='${pic.description ?: "Embedded image"}' />"
        } catch (e: Exception) {
            Logger.e("Error processing embedded picture: ${e.message}", e)
            html += "<!-- Error loading image: ${e.message} -->"
        }
    }
    return html
}

fun convertTableToHtmlExact(table: XWPFTable): String {
    val html = StringBuilder("<table>")
    for (row in table.rows) {
        html.append("<tr>")
        for (cell in row.tableCells) {
            html.append("<td>")
            cell.bodyElements.forEach { element ->
                when (element) {
                    is XWPFParagraph -> html.append(convertParagraphToHtmlExact(element))
                    is XWPFTable -> html.append(convertTableToHtmlExact(element))
                }
            }
            html.append("</td>")
        }
        html.append("</tr>")
    }
    html.append("</table>")
    return html.toString()
}

// =================================================================================
// ENHANCED HTML -> DOCX CONVERSION WITH FULL FORMATTING SUPPORT
// =================================================================================

suspend fun writeHtmlToDocx(context: Context, uri: Uri, htmlContent: String): Result<Unit> {
    Logger.d("Starting enhanced writeHtmlToDocx for $uri")
    return withContext(Dispatchers.IO) {
        runCatching {
            // First, try to preserve the original document structure by reading the existing file
            val originalDocument = try {
                context.contentResolver.openInputStream(uri)?.use { inputStream ->
                    val doc = XWPFDocument(inputStream)
                    // Create a copy of the document structure
                    val newDoc = XWPFDocument()

                    // Copy document properties
                    try {
                        // Document properties are often read-only, so we skip copying them
                        Logger.d("Skipping document properties copying - not supported")
                    } catch (e: Exception) {
                        Logger.w("Could not copy document properties: ${e.message}")
                    }

                    doc.close()
                    newDoc
                }
            } catch (e: Exception) {
                Logger.w("Could not read original document, creating new one: ${e.message}")
                null
            }

            val document = originalDocument ?: XWPFDocument()

            try {
                val rawHtml = cleanHtmlContent(htmlContent)
                Logger.d("Cleaned HTML (first 500 chars): ${rawHtml.take(500)}")

                val jsoupDoc = Jsoup.parse(rawHtml)

                // Set document defaults to ensure proper DOCX structure
                setDocumentDefaults(document)

                // Clear existing content while preserving document structure
                clearDocumentContent(document)

                val body = jsoupDoc.body()
                val docContext = DocumentContext()

                // Parse with enhanced processor
                parseEnhancedJsoupNode(body, document, null, context, docContext)

                // Remove empty first paragraph if exists
                removeEmptyFirstParagraph(document)

                // Ensure document has at least one paragraph (required for valid DOCX)
                ensureMinimumContent(document)

                // Write to temporary buffer first to validate
                val tempBuffer = ByteArrayOutputStream()
                document.write(tempBuffer)

                // If successful, write to actual file
                context.contentResolver.openOutputStream(uri, "wt")?.use { outputStream ->
                    outputStream.write(tempBuffer.toByteArray())
                    outputStream.flush()
                } ?: throw IOException("Unable to open output stream for $uri")

                document.close()
                Logger.i("DOCX saved successfully to $uri with enhanced formatting")
            } catch (e: Exception) {
                document.close()
                throw e
            }
        }.onFailure { e ->
            Logger.e("Error in enhanced writeHtmlToDocx: ${e.message}", e)
        }
    }
}

// Clear document content while preserving structure
private fun clearDocumentContent(document: XWPFDocument) {
    try {
        // Remove all body elements
        val bodyElements = document.bodyElements.toList()
        bodyElements.forEachIndexed { index, _ ->
            if (document.bodyElements.isNotEmpty()) {
                document.removeBodyElement(0)
            }
        }
    } catch (e: Exception) {
        Logger.w("Failed to clear document content: ${e.message}")
    }
}

// Ensure document has minimum required content
private fun ensureMinimumContent(document: XWPFDocument) {
    if (document.bodyElements.isEmpty()) {
        val paragraph = document.createParagraph()
        val run = paragraph.createRun()
        run.setText(" ") // Add a space to ensure the paragraph is not completely empty
    }
}

// Clean HTML content
private fun cleanHtmlContent(htmlContent: String): String {
    var cleaned = htmlContent

    // Remove JSON escaping if present
    if (cleaned.startsWith("\"") && cleaned.endsWith("\"")) {
        cleaned = cleaned.substring(1, cleaned.length - 1)
    }

    cleaned = cleaned
        .replace("\\u003C", "<")
        .replace("\\u003E", ">")
        .replace("\\\"", "\"")
        .replace("\\n", "\n")
        .replace("\\t", "\t")
        .replace("\\r", "\r")
        .replace("&nbsp;", " ")
        .replace("\\\\", "\\")

    // Remove script tags and their content
    cleaned = cleaned.replace(Regex("<script[^>]*>.*?</script>", RegexOption.DOT_MATCHES_ALL), "")

    // Remove style tags but keep the content structure
    cleaned = cleaned.replace(Regex("<style[^>]*>.*?</style>", RegexOption.DOT_MATCHES_ALL), "")

    // Clean up page containers but preserve content
    cleaned = cleaned.replace(Regex("<div[^>]*class=['\"]page['\"][^>]*>"), "<div>")

    return cleaned
}

// Set document defaults with proper validation
private fun setDocumentDefaults(document: XWPFDocument) {
    try {
        val ctDocument = document.document
        if (ctDocument == null) {
            Logger.w("Document CTDocument is null, cannot set defaults")
            return
        }

        val ctBody = ctDocument.body
        if (ctBody == null) {
            Logger.w("Document body is null, cannot set defaults")
            return
        }

        val sectPr = if (ctBody.isSetSectPr) ctBody.sectPr else ctBody.addNewSectPr()

        // Set page size to A4 (210mm x 297mm = 11906 twips x 16838 twips)
        val pgSz = if (sectPr.isSetPgSz) sectPr.pgSz else sectPr.addNewPgSz()
        pgSz.w = BigInteger.valueOf(11906) // A4 width in twips
        pgSz.h = BigInteger.valueOf(16838) // A4 height in twips

        // Set margins (2.54cm = 1440 twips)
        val pgMar = if (sectPr.isSetPgMar) sectPr.pgMar else sectPr.addNewPgMar()
        pgMar.top = BigInteger.valueOf(1440)    // 2.54cm
        pgMar.bottom = BigInteger.valueOf(1440) // 2.54cm
        pgMar.left = BigInteger.valueOf(1440)   // 2.54cm
        pgMar.right = BigInteger.valueOf(1440)  // 2.54cm

        // Set default font for the document

    } catch (e: Exception) {
        Logger.w("Failed to set document defaults: ${e.message}")
    }
}

// Remove empty first paragraph
private fun removeEmptyFirstParagraph(document: XWPFDocument) {
    if (document.paragraphs.isNotEmpty()) {
        val firstPara = document.paragraphs[0]
        if (firstPara.text.isBlank() && firstPara.runs.isEmpty()) {
            if (document.bodyElements.size > 0) {
                document.removeBodyElement(0)
            }
        }
    }
}

// Document context
data class DocumentContext(
    var currentListLevel: Int = 0,
    var isInList: Boolean = false,
    var currentListType: ListType = ListType.BULLET,
    var preserveSpacing: Boolean = true,
    var numberingMap: MutableMap<String, Int> = mutableMapOf()
)

enum class ListType {
    BULLET, NUMBERED
}

// Enhanced HTML node parser with full formatting support
private fun parseEnhancedJsoupNode(
    jsoupNode: Node,
    document: XWPFDocument,
    currentParagraph: XWPFParagraph?,
    context: Context,
    docContext: DocumentContext
): XWPFParagraph? {
    var paragraph = currentParagraph

    when (jsoupNode) {
        is TextNode -> {
            val text = jsoupNode.text().replace("\u00A0", " ")
            if (text.isNotEmpty() || docContext.preserveSpacing) {
                paragraph = paragraph ?: document.createParagraph()
                val run = paragraph.createRun()
                run.setText(text)
            }
        }

        is Element -> {
            val tagName = jsoupNode.tagName().lowercase()

            paragraph = when (tagName) {
                // Block elements
                "p", "div" -> handleParagraphElement(jsoupNode, document, docContext, context)
                "br" -> handleLineBreak(document, paragraph)

                // Headings
                "h1", "h2", "h3", "h4", "h5", "h6" -> handleHeading(jsoupNode, document, tagName, docContext, context)

                // Lists
                "ul" -> handleUnorderedList(jsoupNode, document, docContext, context)
                "ol" -> handleOrderedList(jsoupNode, document, docContext, context)
                "li" -> handleListItem(jsoupNode, document, paragraph, docContext, context)

                // Tables
                "table" -> {
                    handleTable(jsoupNode, document, context, docContext)
                    null
                }

                // Inline formatting
                "span", "strong", "b", "em", "i", "u", "s", "strike", "del",
                "sub", "sup", "code", "mark" -> handleInlineFormatting(jsoupNode, document, paragraph, docContext, context)

                // Images
                "img" -> {
                    handleImage(jsoupNode, document, context)
                    null
                }

                // Blockquote
                "blockquote" -> handleBlockquote(jsoupNode, document, docContext, context)

                // Horizontal rule
                "hr" -> {
                    handleHorizontalRule(document)
                    null
                }

                else -> {
                    // For unknown elements, process children
                    for (child in jsoupNode.childNodes()) {
                        paragraph = parseEnhancedJsoupNode(child, document, paragraph, context, docContext)
                    }
                    paragraph
                }
            }
        }
    }

    return paragraph
}

// Handle paragraph element
private fun handleParagraphElement(
    element: Element,
    document: XWPFDocument,
    docContext: DocumentContext,
    context: Context
): XWPFParagraph {
    val paragraph = document.createParagraph()
    applyParagraphStyle(paragraph, element)

    for (child in element.childNodes()) {
        parseEnhancedJsoupNode(child, document, paragraph, context, docContext)
    }

    return paragraph
}

// Handle line break
private fun handleLineBreak(document: XWPFDocument, currentParagraph: XWPFParagraph?): XWPFParagraph {
    val paragraph = currentParagraph ?: document.createParagraph()
    val run = paragraph.createRun()
    run.addBreak()
    return paragraph
}

// Handle headings
private fun handleHeading(
    element: Element,
    document: XWPFDocument,
    tagName: String,
    docContext: DocumentContext,
    context: Context
): XWPFParagraph {
    val paragraph = document.createParagraph()
    val level = tagName.substring(1).toIntOrNull() ?: 1

    val fontSize = when (level) {
        1 -> 22
        2 -> 18
        3 -> 16
        4 -> 14
        5 -> 12
        else -> 10
    }

    applyParagraphStyle(paragraph, element)

    for (child in element.childNodes()) {
        parseEnhancedJsoupNode(child, document, paragraph, context, docContext)
    }

    // Apply heading formatting to all runs
    paragraph.runs.forEach { run ->
        run.isBold = true
        if (run.fontSize <= 0) {
            run.fontSize = fontSize
        }
    }

    return paragraph
}

// Handle unordered list
private fun handleUnorderedList(
    element: Element,
    document: XWPFDocument,
    docContext: DocumentContext,
    context: Context
): XWPFParagraph? {
    val previousContext = docContext.copy()
    docContext.isInList = true
    docContext.currentListType = ListType.BULLET
    docContext.currentListLevel++

    var lastParagraph: XWPFParagraph? = null
    for (child in element.children()) {
        if (child.tagName().equals("li", ignoreCase = true)) {
            lastParagraph = handleListItem(child, document, null, docContext, context)
        }
    }

    docContext.currentListLevel = previousContext.currentListLevel
    docContext.isInList = previousContext.isInList
    docContext.currentListType = previousContext.currentListType

    return lastParagraph
}

// Handle ordered list
private fun handleOrderedList(
    element: Element,
    document: XWPFDocument,
    docContext: DocumentContext,
    context: Context
): XWPFParagraph? {
    val previousContext = docContext.copy()
    docContext.isInList = true
    docContext.currentListType = ListType.NUMBERED
    docContext.currentListLevel++

    var lastParagraph: XWPFParagraph? = null
    for (child in element.children()) {
        if (child.tagName().equals("li", ignoreCase = true)) {
            lastParagraph = handleListItem(child, document, null, docContext, context)
        }
    }

    docContext.currentListLevel = previousContext.currentListLevel
    docContext.isInList = previousContext.isInList
    docContext.currentListType = previousContext.currentListType

    return lastParagraph
}

// Handle list item
private fun handleListItem(
    element: Element,
    document: XWPFDocument,
    currentParagraph: XWPFParagraph?,
    docContext: DocumentContext,
    context: Context
): XWPFParagraph {
    val paragraph = document.createParagraph()

    // Set numbering
    try {
//        val numId = getOrCreateNumbering(document, docContext)
//        paragraph.numID = BigInteger.valueOf(numId.toLong())
//        paragraph.numILvl = BigInteger.valueOf((docContext.currentListLevel - 1).toLong())
    } catch (e: Exception) {
        Logger.w("Failed to apply numbering: ${e.message}")
        // Fallback: add bullet/number manually
        val run = paragraph.createRun()
        run.setText(if (docContext.currentListType == ListType.BULLET) "• " else "${docContext.currentListLevel}. ")
    }

    // Set indentation
    paragraph.indentationLeft = 720 * docContext.currentListLevel

    for (child in element.childNodes()) {
        parseEnhancedJsoupNode(child, document, paragraph, context, docContext)
    }

    return paragraph
}

// Get or create numbering
//private fun getOrCreateNumbering(document: XWPFDocument, docContext: DocumentContext): Int {
//    val key = "${docContext.currentListType}_${docContext.currentListLevel}"
//
//    return docContext.numberingMap.getOrPut(key) {
//        try {
//            val numbering = document.createNumbering()
//            val abstractNumId = numbering.addAbstractNum()
//            numbering.addNum(abstractNumId)
//        } catch (e: Exception) {
//            1
//        }
//    }
//}

// Handle table with full formatting
private fun handleTable(
    element: Element,
    document: XWPFDocument,
    context: Context,
    docContext: DocumentContext
) {
    try {
        val rows = element.select("tr")
        if (rows.isEmpty()) return

        val firstRow = rows.first()
        val columns = firstRow?.select("td, th")?.size ?: 0

        val table = document.createTable(rows.size, columns)

        // Set table properties
        table.width = 5000

        rows.forEachIndexed { rowIndex, rowElement ->
            val tableRow = table.getRow(rowIndex)
            val cells = rowElement.select("td, th")

            cells.forEachIndexed { colIndex, cellElement ->
                if (colIndex < tableRow.tableCells.size) {
                    val cell = tableRow.getCell(colIndex)

                    // Apply cell styling
                    applyCellStyle(cell, cellElement)

                    // Clear default paragraph
                    if (cell.paragraphs.isNotEmpty()) {
                        cell.removeParagraph(0)
                    }

                    // Parse cell content directly in the cell context
                    var cellParagraph: XWPFParagraph? = null

                    for (child in cellElement.childNodes()) {
                        cellParagraph = parseEnhancedJsoupNodeInCell(
                            child,
                            cell,
                            cellParagraph,
                            context,
                            docContext
                        )
                    }

                    // Make header cells bold
                    if (cellElement.tagName().equals("th", ignoreCase = true)) {
                        cell.paragraphs.forEach { para ->
                            para.runs.forEach { run ->
                                run.isBold = true
                            }
                        }
                    }

                    // Ensure cell has at least one paragraph
                    if (cell.paragraphs.isEmpty()) {
                        cell.addParagraph()
                    }
                }
            }
        }
    } catch (e: Exception) {
        Logger.e("Error creating table: ${e.message}", e)
    }
}

// Enhanced HTML node parser specifically for table cells to handle images correctly
private fun parseEnhancedJsoupNodeInCell(
    jsoupNode: Node,
    cell: XWPFTableCell,
    currentParagraph: XWPFParagraph?,
    context: Context,
    docContext: DocumentContext
): XWPFParagraph? {
    var paragraph = currentParagraph

    when (jsoupNode) {
        is TextNode -> {
            val text = jsoupNode.text().replace("\u00A0", " ")
            if (text.isNotEmpty() || docContext.preserveSpacing) {
                paragraph = paragraph ?: cell.addParagraph()
                val run = paragraph.createRun()
                run.setText(text)
            }
        }

        is Element -> {
            val tagName = jsoupNode.tagName().lowercase()

            paragraph = when (tagName) {
                // Block elements
                "p", "div" -> {
                    val newParagraph = cell.addParagraph()
                    applyParagraphStyle(newParagraph, jsoupNode)

                    for (child in jsoupNode.childNodes()) {
                        paragraph = parseEnhancedJsoupNodeInCell(
                            child,
                            cell,
                            newParagraph,
                            context,
                            docContext
                        )
                    }
                    newParagraph
                }

                "br" -> {
                    val para = paragraph ?: cell.addParagraph()
                    val run = para.createRun()
                    run.addBreak()
                    para
                }

                // Images - handle directly in cell
                "img" -> {
                    val imagePara = paragraph ?: cell.addParagraph()
                    handleImageInCell(jsoupNode, imagePara, context)
                    imagePara
                }

                // Inline formatting
                "span", "strong", "b", "em", "i", "u", "s", "strike", "del",
                "sub", "sup", "code", "mark" -> {
                    val para = paragraph ?: cell.addParagraph()

                    for (child in jsoupNode.childNodes()) {
                        when (child) {
                            is TextNode -> {
                                val run = para.createRun()
                                run.setText(child.text())
                                applyRunFormatting(run, jsoupNode)
                            }

                            is Element -> {
                                paragraph = parseEnhancedJsoupNodeInCell(
                                    child,
                                    cell,
                                    para,
                                    context,
                                    docContext
                                )
                            }
                        }
                    }
                    para
                }

                else -> {
                    // For unknown elements, process children
                    for (child in jsoupNode.childNodes()) {
                        paragraph = parseEnhancedJsoupNodeInCell(
                            child,
                            cell,
                            paragraph,
                            context,
                            docContext
                        )
                    }
                    paragraph
                }
            }
        }
    }

    return paragraph
}

// Handle image specifically within a table cell paragraph
private fun handleImageInCell(element: Element, paragraph: XWPFParagraph, context: Context) {
    try {
        val src = element.attr("src")

        if (src.startsWith("data:image")) {
            val parts = src.split(",")
            if (parts.size == 2) {
                val meta = parts[0]
                val base64Data = parts[1]
                val decodedData = Base64.decode(base64Data, Base64.DEFAULT)

                val pictureType = when {
                    meta.contains("jpeg") -> XWPFDocument.PICTURE_TYPE_JPEG
                    meta.contains("png") -> XWPFDocument.PICTURE_TYPE_PNG
                    meta.contains("gif") -> XWPFDocument.PICTURE_TYPE_GIF
                    meta.contains("bmp") -> XWPFDocument.PICTURE_TYPE_BMP
                    meta.contains("tiff") -> XWPFDocument.PICTURE_TYPE_TIFF
                    else -> XWPFDocument.PICTURE_TYPE_PNG
                }

                val widthPx = element.attr("width").removeSuffix("px").toIntOrNull() ?: 200
                val heightPx = element.attr("height").removeSuffix("px").toIntOrNull() ?: 150
                val widthEmu = widthPx * Units.EMU_PER_PIXEL
                val heightEmu = heightPx * Units.EMU_PER_PIXEL

                paragraph.createRun().addPicture(
                    ByteArrayInputStream(decodedData),
                    pictureType,
                    "image.dat",
                    widthEmu,
                    heightEmu
                )
            }
        }
    } catch (e: Exception) {
        Logger.e("Error processing image in cell: ${e.message}", e)
    }
}

// Handle inline formatting
private fun handleInlineFormatting(
    element: Element,
    document: XWPFDocument,
    currentParagraph: XWPFParagraph?,
    docContext: DocumentContext,
    context: Context
): XWPFParagraph {
    val paragraph = currentParagraph ?: document.createParagraph()

    for (child in element.childNodes()) {
        when (child) {
            is TextNode -> {
                val run = paragraph.createRun()
                run.setText(child.text())
                applyRunFormatting(run, element)
            }
            is Element -> {
                parseEnhancedJsoupNode(child, document, paragraph, context, docContext)
            }
        }
    }

    // Apply formatting to runs that were just created
    val tagName = element.tagName().lowercase()
    if (tagName in listOf("span", "strong", "b", "em", "i", "u", "s", "strike", "del", "sub", "sup", "code", "mark")) {
        // Get recently created runs for this element
        val recentRunCount = element.childNodes().filterIsInstance<TextNode>().size
        if (recentRunCount > 0 && paragraph.runs.size >= recentRunCount) {
            val runsToFormat = paragraph.runs.takeLast(recentRunCount)
            runsToFormat.forEach { run ->
                applyRunFormatting(run, element)
            }
        }
    }

    return paragraph
}

// Apply run formatting
private fun applyRunFormatting(run: XWPFRun, element: Element) {
    val tagName = element.tagName().lowercase()
    val style = element.attr("style")

    when (tagName) {
        "strong", "b" -> run.isBold = true
        "em", "i" -> run.isItalic = true
        "u" -> run.underline = UnderlinePatterns.SINGLE
        "s", "strike", "del" -> run.isStrikeThrough = true
//        "sub" -> run.subscript = VerticalAlign.SUBSCRIPT
//        "sup" -> run.subscript = VerticalAlign.SUPERSCRIPT
        "code" -> {
            run.fontFamily = "Courier New"
            if (run.fontSize <= 0) run.fontSize = 10
        }
        "mark" -> {
            try {
                run.setTextHighlightColor("yellow")
            } catch (e: Exception) {
                Logger.w("Failed to set highlight: ${e.message}")
            }
        }
    }

    // Parse inline styles
    if (style.isNotBlank()) {
        applyInlineStyle(run, style)
    }

    // Apply color attribute
    val color = element.attr("color")
    if (color.isNotBlank()) {
        run.color = color.replace("#", "")
    }
}

// Apply inline CSS styles
private fun applyInlineStyle(run: XWPFRun, style: String) {
    val styles = style.split(";").associate {
        val parts = it.split(":")
        if (parts.size == 2) {
            parts[0].trim().lowercase() to parts[1].trim()
        } else {
            "" to ""
        }
    }

    styles["font-weight"]?.let {
        if (it.contains("bold", ignoreCase = true) || it.toIntOrNull()?.let { w -> w >= 600 } == true) {
            run.isBold = true
        }
    }

    styles["font-style"]?.let {
        if (it.contains("italic", ignoreCase = true)) {
            run.isItalic = true
        }
    }

    styles["text-decoration"]?.let {
        when {
            it.contains("underline", ignoreCase = true) -> run.underline = UnderlinePatterns.SINGLE
            it.contains("line-through", ignoreCase = true) -> run.isStrikeThrough = true
        }
    }

    styles["font-size"]?.let {
        val size = when {
            it.endsWith("pt") -> it.removeSuffix("pt").toDoubleOrNull()?.toInt()
            it.endsWith("px") -> (it.removeSuffix("px").toDoubleOrNull()?.times(0.75))?.toInt()
            else -> null
        }
        size?.let { s -> run.fontSize = s }
    }

    styles["font-family"]?.let {
        run.fontFamily = it.replace("'", "").replace("\"", "").split(",")[0].trim()
    }

    styles["color"]?.let {
        val colorValue = it.replace("#", "").trim()
        if (colorValue.length == 6 || colorValue.length == 3) {
            run.color = colorValue
        }
    }
}

// Handle image
private fun handleImage(element: Element, document: XWPFDocument, context: Context) {
    try {
        val src = element.attr("src")

        if (src.startsWith("data:image")) {
            val parts = src.split(",")
            if (parts.size == 2) {
                val meta = parts[0]
                val base64Data = parts[1]
                val decodedData = Base64.decode(base64Data, Base64.DEFAULT)

                val pictureType = when {
                    meta.contains("jpeg") -> XWPFDocument.PICTURE_TYPE_JPEG
                    meta.contains("png") -> XWPFDocument.PICTURE_TYPE_PNG
                    meta.contains("gif") -> XWPFDocument.PICTURE_TYPE_GIF
                    meta.contains("bmp") -> XWPFDocument.PICTURE_TYPE_BMP
                    meta.contains("tiff") -> XWPFDocument.PICTURE_TYPE_TIFF
                    else -> XWPFDocument.PICTURE_TYPE_PNG
                }

                val widthPx = element.attr("width").removeSuffix("px").toIntOrNull() ?: 200
                val heightPx = element.attr("height").removeSuffix("px").toIntOrNull() ?: 150
                val widthEmu = widthPx * Units.EMU_PER_PIXEL
                val heightEmu = heightPx * Units.EMU_PER_PIXEL

                val imageParagraph = document.createParagraph()
                imageParagraph.createRun().addPicture(
                    ByteArrayInputStream(decodedData),
                    pictureType,
                    "image.dat",
                    widthEmu,
                    heightEmu
                )
            }
        }
    } catch (e: Exception) {
        Logger.e("Error processing image: ${e.message}", e)
    }
}

// Handle blockquote
private fun handleBlockquote(
    element: Element,
    document: XWPFDocument,
    docContext: DocumentContext,
    context: Context
): XWPFParagraph? {
    var lastParagraph: XWPFParagraph? = null

    for (child in element.childNodes()) {
        lastParagraph = parseEnhancedJsoupNode(child, document, lastParagraph, context, docContext)
        lastParagraph?.let { para ->
            para.indentationLeft = 720 // Indent blockquote
            para.borderLeft = Borders.SINGLE
        }
    }

    return lastParagraph
}

// Handle horizontal rule
private fun handleHorizontalRule(document: XWPFDocument): XWPFParagraph {
    val paragraph = document.createParagraph()
    paragraph.borderBottom = Borders.SINGLE
    return paragraph
}

// Apply paragraph style from HTML element
private fun applyParagraphStyle(paragraph: XWPFParagraph, element: Element) {
    val style = element.attr("style")

    if (style.isNotBlank()) {
        val styles = style.split(";").associate {
            val parts = it.split(":")
            if (parts.size == 2) {
                parts[0].trim().lowercase() to parts[1].trim()
            } else {
                "" to ""
            }
        }

        styles["text-align"]?.let {
            paragraph.alignment = when (it.lowercase()) {
                "center" -> ParagraphAlignment.CENTER
                "right" -> ParagraphAlignment.RIGHT
                "justify" -> ParagraphAlignment.BOTH
                "left" -> ParagraphAlignment.LEFT
                else -> paragraph.alignment
            }
        }

        styles["margin-top"]?.let {
            val spacing = when {
                it.endsWith("pt") -> it.removeSuffix("pt").toDoubleOrNull()?.times(20)?.toInt()
                it.endsWith("px") -> it.removeSuffix("px").toDoubleOrNull()?.times(15)?.toInt()
                else -> null
            }
            spacing?.let { s -> paragraph.spacingBefore = s }
        }

        styles["margin-bottom"]?.let {
            val spacing = when {
                it.endsWith("pt") -> it.removeSuffix("pt").toDoubleOrNull()?.times(20)?.toInt()
                it.endsWith("px") -> it.removeSuffix("px").toDoubleOrNull()?.times(15)?.toInt()
                else -> null
            }
            spacing?.let { s -> paragraph.spacingAfter = s }
        }

        styles["line-height"]?.let {
            val lineHeightValue = it.toDoubleOrNull()
            lineHeightValue?.let { lh ->
                paragraph.setSpacingBetween(lh * 240, LineSpacingRule.AUTO)
            }
        }
    }

    // Apply alignment from attribute
    val align = element.attr("align")
    if (align.isNotBlank()) {
        paragraph.alignment = when (align.lowercase()) {
            "center" -> ParagraphAlignment.CENTER
            "right" -> ParagraphAlignment.RIGHT
            "justify" -> ParagraphAlignment.BOTH
            "left" -> ParagraphAlignment.LEFT
            else -> paragraph.alignment
        }
    }
}

// Apply cell style
private fun applyCellStyle(cell: XWPFTableCell, element: Element) {
    try {
        // Background color
        val bgColor = element.attr("bgcolor")
        if (bgColor.isNotBlank()) {
            cell.color = bgColor.replace("#", "")
        }

        // Parse style attribute for background color
        val style = element.attr("style")
        if (style.contains("background-color")) {
            val bgMatch = Regex("background-color:\\s*#?([0-9A-Fa-f]{6})").find(style)
            bgMatch?.let {
                cell.color = it.groupValues[1]
            }
        }

        // Set cell margins
        val ctTc = cell.ctTc
        val tcPr = ctTc.tcPr ?: ctTc.addNewTcPr()
        val tcMar = tcPr.tcMar ?: tcPr.addNewTcMar()

        tcMar.left = CTTblWidth.Factory.newInstance().apply {
            w = BigInteger.valueOf(100)
            type = STTblWidth.DXA
        }
        tcMar.right = CTTblWidth.Factory.newInstance().apply {
            w = BigInteger.valueOf(100)
            type = STTblWidth.DXA
        }
    } catch (e: Exception) {
        Logger.w("Failed to apply cell style: ${e.message}")
    }
}

// =================================================================================
// LEGACY FUNCTIONS (Keep for backward compatibility)
// =================================================================================

private fun parseJsoupNode(
    jsoupNode: Node,
    document: XWPFDocument,
    currentParagraph: XWPFParagraph?,
    context: Context
) {
    // Use enhanced parser
    val docContext = DocumentContext()
    parseEnhancedJsoupNode(jsoupNode, document, currentParagraph, context, docContext)
}

private fun applyParagraphAlignment(element: Element, paragraph: XWPFParagraph?) {
    if (paragraph == null) return
    applyParagraphStyle(paragraph, element)
}

fun parseStyleAndApply(styleAttribute: String, paragraph: XWPFParagraph?, run: XWPFRun?) {
    if (run == null && paragraph == null) return

    styleAttribute.split(';').map { it.trim() }.filter { it.isNotEmpty() }.forEach { style ->
        val parts = style.split(':').map { it.trim() }
        if (parts.size == 2) {
            val property = parts[0].lowercase()
            val value = parts[1]
            try {
                when (property) {
                    // Run-specific styles
                    "color" -> if (value.startsWith("#") && (value.length == 7 || value.length == 4)) {
                        run?.setColor(value.removePrefix("#"))
                    }
                    "font-weight" -> if (value == "bold" || (value.toIntOrNull() ?: 400) >= 600) {
                        run?.isBold = true
                    }
                    "font-style" -> if (value == "italic") {
                        run?.isItalic = true
                    }
                    "font-size" -> {
                        val size = when {
                            value.endsWith("pt") -> value.removeSuffix("pt").toDoubleOrNull()?.toInt()
                            value.endsWith("px") -> (value.removeSuffix("px").toDoubleOrNull()?.times(0.75))?.toInt()
                            else -> null
                        }
                        if (size != null) run?.fontSize = size
                    }
                    "font-family" -> run?.fontFamily = value.replace("\'", "").split(",")[0].trim()
                    "text-decoration", "text-decoration-line" -> {
                        if (value.contains("underline")) run?.underline = UnderlinePatterns.SINGLE
                        if (value.contains("line-through")) run?.isStrikeThrough = true
                    }
                    // Paragraph-specific styles
                    "margin-top" -> {
                        if (paragraph != null && value.endsWith("pt")) {
                            paragraph.spacingBefore = (value.removeSuffix("pt").toDoubleOrNull()?.times(20))?.toInt() ?: paragraph.spacingBefore
                        }
                    }
                    "margin-bottom" -> {
                        if (paragraph != null && value.endsWith("pt")) {
                            paragraph.spacingAfter = (value.removeSuffix("pt").toDoubleOrNull()?.times(20))?.toInt() ?: paragraph.spacingAfter
                        }
                    }
                    "line-height" -> {
                        if (paragraph != null) {
                            val lineHeightValue = value.toDoubleOrNull()
                            if (lineHeightValue != null) {
                                paragraph.setSpacingBetween(lineHeightValue * 240, LineSpacingRule.AUTO)
                            }
                        }
                    }
                }
            } catch (e: Exception) {
                Logger.w("Failed to parse style: $property = $value. Error: ${e.message}")
            }
        }
    }
}