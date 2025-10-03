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
//import androidx.compose.material.icons.filled.FormatItalic // Ensured active import
//import androidx.compose.material.icons.filled.FormatUnderlined // Ensured active import
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
import androidx.navigation.NavController
import com.example.docx.R // For custom drawable R.drawable.bold
import com.example.docx.util.Logger
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import org.apache.poi.common.usermodel.PictureType
import org.apache.poi.util.Units
import org.apache.poi.xwpf.usermodel.*
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STVerticalAlignRun as OfficeSTVerticalAlignRun
import org.jsoup.Jsoup
import org.jsoup.nodes.Element
import org.jsoup.nodes.Node
import org.jsoup.nodes.TextNode
import java.io.ByteArrayInputStream
import java.io.IOException

@OptIn(ExperimentalMaterial3Api::class)
@SuppressLint("SetJavaScriptEnabled")
@Composable
fun CreateNewDocxScreen(navController: NavController) {
    var documentHtml by remember {
        mutableStateOf(
            """
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
                body {
                    font-family: 'Times New Roman', serif;
                    line-height: 1.6;
                    background: #f0f0f0;
                    margin: 0;
                    padding: 20px;
                }
                .page {
                    background: white;
                    margin: 0 auto 20px auto;
                    padding-top: 1.00cm;
                    padding-bottom: 1.00cm;
                    padding-left: 2.54cm;
                    padding-right: 2.54cm;
                    width: 21cm;
                    min-height: 29.7cm;
                    box-shadow: 0 4px 8px rgba(0,0,0,0.15);
                    page-break-after: always;
                    position: relative;
                    overflow: visible;
                }
                .page:last-child { page-break-after: auto; }
                .page-number {
                    position: absolute;
                    bottom: 1.27cm;
                    right: 2.54cm;
                    font-size: 10pt;
                    color: #666;
                }
                h1, h2, h3, h4, h5 { margin-top: 0.4em; margin-bottom: 0.2em; }
                p { margin: 0.2em 0; }
                ul, ol { margin: 0.5em 0 0.5em 1em; }
                table { border-collapse: collapse; width: 100%; margin: 1em 0; }
                td, th { border: 1px solid #444; padding: 3px; text-align: left; }
                img { max-width: 100%; height: auto; display: block; margin: 0.5em auto; }
            </style>
            <script>
                let pageCounter = 1;
                let isProcessing = false;
                
                function createNewPage() {
                    pageCounter++;
                    const newPage = document.createElement('div');
                    newPage.className = 'page';
                    newPage.id = 'page' + pageCounter;
                    newPage.innerHTML = '<p><br></p>';
                    
                    // Add page number
                    const pageNum = document.createElement('div');
                    pageNum.className = 'page-number';
                    pageNum.style.cssText = 'position: absolute; bottom: 1.27cm; right: 2.54cm; font-size: 10pt; color: #666; pointer-events: none;';
                    pageNum.textContent = pageCounter;
                    newPage.appendChild(pageNum);
                    
                    document.body.appendChild(newPage);
                    return newPage;
                }
                
                function updatePageNumbers() {
                    const pages = document.querySelectorAll('.page');
                    pages.forEach((page, index) => {
                        let pageNum = page.querySelector('.page-number');
                        if (!pageNum) {
                            pageNum = document.createElement('div');
                            pageNum.className = 'page-number';
                            pageNum.style.cssText = 'position: absolute; bottom: 1.27cm; right: 2.54cm; font-size: 10pt; color: #666; pointer-events: none;';
                            page.appendChild(pageNum);
                        }
                        pageNum.textContent = index + 1;
                    });
                }
                
                function checkPageOverflow() {
                    if (isProcessing) return;
                    isProcessing = true;
                    
                    const pages = document.querySelectorAll('.page');
                    const pageHeight = 29.7 * 37.795; // A4 height in pixels
                    
                    pages.forEach((page, index) => {
                        const contentHeight = page.scrollHeight;
                        if (contentHeight > pageHeight) {
                            // Create new page if it doesn't exist
                            let nextPage = document.getElementById('page' + (index + 2));
                            if (!nextPage) {
                                nextPage = createNewPage();
                            }
                            
                            // Move overflow content to next page
                            const elements = Array.from(page.children);
                            let totalHeight = 0;
                            let moveIndex = -1;
                            
                            for (let i = 0; i < elements.length; i++) {
                                const elem = elements[i];
                                if (elem.className === 'page-number') continue;
                                
                                const elemHeight = elem.offsetHeight;
                                if (totalHeight + elemHeight > pageHeight - 100) { // Leave some margin
                                    moveIndex = i;
                                    break;
                                }
                                totalHeight += elemHeight;
                            }
                            
                            if (moveIndex !== -1) {
                                for (let i = moveIndex; i < elements.length; i++) {
                                    const elem = elements[i];
                                    if (elem.className !== 'page-number') {
                                        nextPage.insertBefore(elem, nextPage.querySelector('.page-number'));
                                    }
                                }
                            }
                        }
                    });
                    
                    updatePageNumbers();
                    isProcessing = false;
                }
                
                function ensureMinimumOnePage() {
                    if (document.querySelectorAll('.page').length === 0) {
                        const firstPage = document.createElement('div');
                        firstPage.className = 'page';
                        firstPage.id = 'page1';
                        firstPage.innerHTML = '<p><br></p>';
                        document.body.appendChild(firstPage);
                    }
                    updatePageNumbers();
                }
                
                document.addEventListener('DOMContentLoaded', function() {
                    ensureMinimumOnePage();
                    setTimeout(checkPageOverflow, 500);
                });
                
                document.addEventListener('input', function() {
                    setTimeout(checkPageOverflow, 300);
                });
                
                document.addEventListener('keydown', function(e) {
                    if (e.ctrlKey && e.key === 'Enter') {
                        e.preventDefault();
                        createNewPage();
                        updatePageNumbers();
                    }
                });
                
                // Periodic check for content changes
                setInterval(checkPageOverflow, 2000);
            </script>
        </head>
        <body>
            <div class="page" id="page1">
                <p><br></p>
            </div>
        </body>
        </html>
    """.trimIndent()
        )
    }
    var isSaving by remember { mutableStateOf(false) }
    val context = LocalContext.current
    val scope = rememberCoroutineScope()
    var errorMessage by remember { mutableStateOf<String?>(null) }
    var showErrorDialog by remember { mutableStateOf(false) }
    val webView = remember { WebView(context) }

    val createDocumentLauncher = rememberLauncherForActivityResult(
        contract = ActivityResultContracts.CreateDocument("application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
        onResult = { newFileUri: Uri? ->
            if (newFileUri != null) {
                webView.evaluateJavascript(
                    "(function() { return document.documentElement.outerHTML; })();"
                ) { htmlContent ->
                    scope.launch {
                        isSaving = true
                        val result = writeHtmlToDoc(context, newFileUri, htmlContent ?: "<html><body><p></p></body></html>") // Corrected name
                        result.fold(
                            onSuccess = {
                                Logger.i("Document saved successfully to $newFileUri from CreateNewDocxScreen")
                                navController.popBackStack()
                            },
                            onFailure = { error ->
                                errorMessage = "Save failed: ${error.message}"
                                showErrorDialog = true
                            }
                        )
                        if (!showErrorDialog) isSaving = false
                    }
                }
            } else {
                isSaving = false
                Logger.d("Save As action cancelled by user in CreateNewDocxScreen.")
            }
        }
    )

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
            FormattingToolbartop(webView = webView)
            Box(modifier = Modifier.weight(1f)) {
                AndroidView(
                    factory = { webViewContext ->
                        webView.apply {
                            settings.javaScriptEnabled = true
                            settings.domStorageEnabled = true
                            settings.allowFileAccess = true
                            settings.defaultTextEncodingName = "utf-8"
                            settings.builtInZoomControls = true
                            settings.displayZoomControls = false
                            settings.setSupportZoom(true)

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
                        .padding(horizontal = 16.dp),
                    update = { /* No specific update logic needed here for now */ }
                )
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
fun FormattingToolbartop(webView: WebView) {
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
                painter = painterResource(R.drawable.bold), // Corrected icon
                contentDescription = "Italic",
                tint = if (isItalicSelected) MaterialTheme.colorScheme.primary else Color.Unspecified
            )
        }

        IconButton(onClick = {
            isUnderlineSelected = !isUnderlineSelected
            webView.evaluateJavascript("document.execCommand('underline');", null)
        }) {
            Icon(
                painter = painterResource(R.drawable.bold), // Corrected icon
                contentDescription = "Underline",
                tint = if (isUnderlineSelected) MaterialTheme.colorScheme.primary else Color.Unspecified
            )
        }
    }
}

// DOCX -> HTML conversion functions (These are defined but not actively used for DOCX loading in this screen)
fun convertParagraphToHtml(paragraph: XWPFParagraph): String {
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
        val runHtml = convertRunToHtml(run)
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

fun convertRunToHtml(run: XWPFRun): String {
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
            Logger.e("Error processing embedded picture in CreateNewDocxScreen: ${e.message}", e)
            html += "<!-- Error loading image: ${e.message} -->"
        }
    }
    return html
}

fun convertTableToHtml(table: XWPFTable): String {
    val html = StringBuilder("<table>")
    for (row in table.rows) {
        html.append("<tr>")
        for (cell in row.tableCells) {
            html.append("<td>")
            cell.bodyElements.forEach { element ->
                when (element) {
                    is XWPFParagraph -> html.append(convertParagraphToHtml(element))
                    is XWPFTable -> html.append(convertTableToHtml(element))
                }
            }
            html.append("</td>")
        }
        html.append("</tr>")
    }
    html.append("</table>")
    return html.toString()
}

// HTML -> DOCX CONVERSION
suspend fun writeHtmlToDoc(context: Context, uri: Uri, htmlContent: String): Result<Unit> {
    Logger.d("Starting writeHtmlToDoc for $uri in CreateNewDocxScreen")
    return withContext(Dispatchers.IO) {
        runCatching {
            val rawHtml = htmlContent.let {
                var h = it
                if (h.startsWith("\"") && h.endsWith("\"")) {
                    h = h.substring(1, h.length - 1)
                }
                h.replace("\\u003C", "<")
                 .replace("\\\"", "\"")
                 .replace("\\n", "\n")
            }
            Logger.d("Cleaned HTML (first 500 chars) in CreateNewDocxScreen: ${rawHtml.take(500)}")
            val jsoupDoc: org.jsoup.nodes.Document = Jsoup.parse(rawHtml)

            context.contentResolver.openOutputStream(uri)?.use { outputStream ->
                val document = XWPFDocument()
                // Set A4 page size and margins
                setA4PageSize(document)
                val body = jsoupDoc.body()
                parseJsoupNode(body, document, null, context)

                if (document.paragraphs.isNotEmpty() && document.paragraphs[0].text.isBlank() && document.paragraphs[0].runs.isEmpty()) {
                    if(document.bodyElements.size > 0) document.removeBodyElement(0)
                }
                document.write(outputStream)
                document.close()
                Logger.i("DOCX saved successfully to $uri from CreateNewDocxScreen")
            } ?: throw IOException("Unable to open output stream for $uri")
        }.onFailure { e ->
            Logger.e("Error in writeHtmlToDoc in CreateNewDocxScreen: ${e.message}", e)
        }
    }
}

// Set A4 page size and margins
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

private fun parseJsoupNode(
    jsoupNode: Node,
    document: XWPFDocument,
    currentParagraph: XWPFParagraph?,
    context: Context
) {
    when (jsoupNode) {
        is TextNode -> {
            val text = jsoupNode.text().replace("\u00A0", " ")
            if (text.isNotEmpty()) {
                val para = currentParagraph ?: document.createParagraph()
                para.createRun().setText(text)
            }
        }
        is Element -> {
            var nextParagraph: XWPFParagraph? = currentParagraph
            var currentRun: XWPFRun? = null

            val tagName = jsoupNode.tagName().lowercase()
            val isBlockElement = tagName in listOf("p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "table", "ul", "ol", "li", "blockquote", "hr", "header", "footer", "section", "article", "aside", "nav")

            if (isBlockElement || nextParagraph == null) {
                nextParagraph = document.createParagraph()
            }
            currentRun = nextParagraph.createRun()

            parseStyleAndApply(jsoupNode.attr("style"), nextParagraph, currentRun)
            applyParagraphAlignment(jsoupNode, nextParagraph)

            when (tagName) {
                "h1" -> { currentRun?.fontSize = 22; currentRun?.isBold = true }
                "h2" -> { currentRun?.fontSize = 18; currentRun?.isBold = true }
                "h3" -> { currentRun?.fontSize = 16; currentRun?.isBold = true }
                "h4" -> { currentRun?.fontSize = 14; currentRun?.isBold = true }
                "h5" -> { currentRun?.fontSize = 12; currentRun?.isBold = true }
                "h6" -> { currentRun?.fontSize = 10; currentRun?.isBold = true }
                "strong", "b" -> currentRun?.isBold = true
                "em", "i" -> currentRun?.isItalic = true
                "u" -> currentRun?.underline = UnderlinePatterns.SINGLE
                "s", "strike" -> currentRun?.isStrikeThrough = true
                "br" -> currentRun?.addBreak()
                "img" -> {
                    val src = jsoupNode.attr("src")
                    if (src.startsWith("data:image")) {
                        try {
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
                                val widthPx = jsoupNode.attr("width").removeSuffix("px").toIntOrNull() ?: 200
                                val heightPx = jsoupNode.attr("height").removeSuffix("px").toIntOrNull() ?: 150
                                val widthEmu = widthPx * Units.EMU_PER_PIXEL
                                val heightEmu = heightPx * Units.EMU_PER_PIXEL

                                val imageParagraph = document.createParagraph()
                                imageParagraph.createRun().addPicture(ByteArrayInputStream(decodedData), pictureType, "image.dat", widthEmu, heightEmu)
                                nextParagraph = null 
                            }
                        } catch (e: Exception) {
                            Logger.e("Error processing base64 image in CreateNewDocxScreen: ${e.message}", e)
                        }
                    }
                }
            }

            val paragraphForChildren = if (tagName == "img") null else nextParagraph
            jsoupNode.childNodes().forEach { child ->
                parseJsoupNode(child, document, paragraphForChildren, context)
            }
        }
    }
}

private fun applyParagraphAlignment(element: Element, paragraph: XWPFParagraph?) {
    if (paragraph == null) return
    val style = element.attr("style")
    when {
        style.contains("text-align: center") -> paragraph.alignment = ParagraphAlignment.CENTER
        style.contains("text-align: right") -> paragraph.alignment = ParagraphAlignment.RIGHT
        style.contains("text-align: justify") -> paragraph.alignment = ParagraphAlignment.BOTH
        style.contains("text-align: left") -> paragraph.alignment = ParagraphAlignment.LEFT
        else -> {
            val classes = element.className()
            when {
                classes.contains("text-center") -> paragraph.alignment = ParagraphAlignment.CENTER
                classes.contains("text-right") -> paragraph.alignment = ParagraphAlignment.RIGHT
                classes.contains("text-justify") -> paragraph.alignment = ParagraphAlignment.BOTH
                classes.contains("text-left") -> paragraph.alignment = ParagraphAlignment.LEFT
            }
        }
    }
}

//private fun parseStyleAndApply(styleAttribute: String, paragraph: XWPFParagraph?, run: XWPFRun?) {
//    if (run == null && paragraph == null) return
//
//    styleAttribute.split(';').map { it.trim() }.filter { it.isNotEmpty() }.forEach { style ->
//        val parts = style.split(':').map { it.trim() }
//        if (parts.size == 2) {
//            val property = parts[0].lowercase()
//            val value = parts[1]
//            try {
//                when (property) {
//                    "color" -> if (value.startsWith("#") && (value.length == 7 || value.length == 4)) {
//                        run?.setColor(value.removePrefix("#"))
//                    }
//                    "font-weight" -> if (value == "bold" || (value.toIntOrNull() ?: 400) >= 600) {
//                        run?.isBold = true
//                    }
//                    "font-style" -> if (value == "italic") {
//                        run?.isItalic = true
//                    }
//                    "font-size" -> {
//                        val size = when {
//                            value.endsWith("pt") -> value.removeSuffix("pt").toDoubleOrNull()?.toInt()
//                            value.endsWith("px") -> (value.removeSuffix("px").toDoubleOrNull()?.times(0.75))?.toInt()
//                            else -> null
//                        }
//                        if (size != null) run?.fontSize = size
//                    }
//                    "font-family" -> run?.fontFamily = value.replace("\'", "").split(",")[0].trim()
//                    "text-decoration", "text-decoration-line" -> {
//                        if (value.contains("underline")) run?.underline = UnderlinePatterns.SINGLE
//                        if (value.contains("line-through")) run?.isStrikeThrough = true
//                    }
//                    "margin-top" -> {
//                        if (paragraph != null && value.endsWith("pt")) {
//                            paragraph.spacingBefore = (value.removeSuffix("pt").toDoubleOrNull()?.times(20))?.toInt() ?: paragraph.spacingBefore
//                        }
//                    }
//                    "margin-bottom" -> {
//                        if (paragraph != null && value.endsWith("pt")) {
//                            paragraph.spacingAfter = (value.removeSuffix("pt").toDoubleOrNull()?.times(20))?.toInt() ?: paragraph.spacingAfter
//                        }
//                    }
//                    "line-height" -> {
//                        if (paragraph != null) {
//                            val lineHeightValue = value.toDoubleOrNull()
//                            if (lineHeightValue != null) {
//                                paragraph.setSpacingBetween(lineHeightValue * 240, LineSpacingRule.AUTO)
//                            }
//                        }
//                    }
//                }
//            } catch (e: Exception) {
//                Logger.w("Failed to parse style in CreateNewDocxScreen: $property = $value. Error: ${e.message}")
//            }
//        }
//    }
//}
