package com.example.docx.ui

import android.graphics.Color
import android.net.Uri
import android.webkit.WebView
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.padding
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Edit
import androidx.compose.material3.*
import androidx.compose.runtime.Composable
import androidx.compose.runtime.remember
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.viewinterop.AndroidView
import androidx.documentfile.provider.DocumentFile
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLConverter
import fr.opensagres.poi.xwpf.converter.xhtml.XHTMLOptions
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.ByteArrayOutputStream
import java.io.InputStream

// --- Convert DOCX -> HTML ---
fun convertDocxToHtml(inputStream: InputStream): String {
    val document = XWPFDocument(inputStream)
    val out = ByteArrayOutputStream()
    val options = XHTMLOptions.create().apply {
        // You can configure options here (e.g. images handling)
    }
    XHTMLConverter.getInstance().convert(document, out, options)
    return out.toString("UTF-8")
}

// --- Composable to show DOCX ---
@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun ViewDocxScreen(
    htmlContent: String,
    fileUriString: String,
    onNavigateBack: () -> Unit,
    onNavigateToEdit: (String) -> Unit,

) {
    val context = LocalContext.current
    val documentName = remember(fileUriString) {
        DocumentFile.fromSingleUri(context, Uri.parse(fileUriString))?.name ?: "Document"
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = { Text(documentName) },
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
        }
        ,modifier = Modifier.fillMaxSize()
    ) { paddingValues ->
        AndroidView(
            factory = { ctx ->
                WebView(ctx).apply {
                    settings.javaScriptEnabled = true
                    settings.allowFileAccess = false
                    settings.setSupportZoom(true)
                    settings.builtInZoomControls = true
                    settings.displayZoomControls = false
                    setBackgroundColor(Color.WHITE)
                }
            },
            update = { webView ->
                val styledHtml = """
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
//                            padding: 2.54cm; /* A4 standard margins */
                            padding-top: 1.00cm; /* A4 standard margins */
                            padding-bottom: 1.00cm; /* A4 standard margins */
                            padding-left: 2.54cm; /* A4 standard margins */
                            padding-right: 2.54cm; /* A4 standard margins */
                            width: 21cm;   /* A4 width */
                            min-height: 29.7cm; /* A4 height */
                            box-shadow: 0 4px 8px rgba(0,0,0,0.15);
                            page-break-after: always;
                            position: relative;
                            overflow: visible;
                        }
                        .page:last-child {
                            page-break-after: auto;
                        }
                        .page-number {
                            position: absolute;
                            bottom: 1.27cm;
                            right: 2.54cm;
                            font-size: 10pt;
                            color: #666;
                        }

                        h1, h2, h3, h4, h5 {
                            margin-top: 0.4em;
                            margin-bottom: 0.2em;
                        }
                        p {
                            margin: 0.2em 0;
                        }
                        ul, ol {
                            margin: 0.5em 0 0.5em 1em;
                        }
                        table {
                            border-collapse: collapse;
                            width: 100%;
                            margin: 1em 0;
                        }
                        td, th {
                            border: 1px solid #444;
                            padding: 3px;
                            text-align: left;
                        }
                        img {
                            max-width: 100%;
                            height: auto;
                            display: block;
                            margin: 0.5em auto;
                        }
                      </style>
                      <script>
                        function splitContentIntoPages(content) {
                            const tempDiv = document.createElement('div');
                            tempDiv.innerHTML = content;
                            tempDiv.style.cssText = 'position: absolute; top: -10000px; width: 21cm;padding: 2.54cm;font-family: Times New Roman; line-height: 1.6;'; 
                            document.body.appendChild(tempDiv);
                            
                            const pageHeight = 29.7 * 37.795; // A4 height in pixels
                            const pages = [];
                            let currentPageContent = '';
                            let currentHeight = 0;
                            let pageNumber = 1;
                            
                            const elements = Array.from(tempDiv.children);
                            
                            for (const element of elements) {
                                const elementHeight = element.offsetHeight;
                                
                                if (currentHeight + elementHeight > pageHeight - 200) { // Leave margin for page number and padding
                                    // Finish current page
                                    if (currentPageContent) {
                                        pages.push({
                                            content: currentPageContent,
                                            number: pageNumber++
                                        });
                                        currentPageContent = '';
                                        currentHeight = 0;
                                    }
                                }
                                
                                currentPageContent += element.outerHTML;
                                currentHeight += elementHeight;
                            }
                            
                            // Add remaining content as last page
                            if (currentPageContent) {
                                pages.push({
                                    content: currentPageContent,
                                    number: pageNumber
                                });
                            }
                            
                            // Ensure at least one page
                            if (pages.length === 0) {
                                pages.push({
                                    content: content || '<p><br></p>',
                                    number: 1
                                });
                            }
                            
                            document.body.removeChild(tempDiv);
                            return pages;
                        }
                        
                        function renderPages() {
                            const container = document.querySelector('.page');
                            if (!container) return;
                            
                            const originalContent = container.innerHTML;
                            const pages = splitContentIntoPages(originalContent);
                            
                            // Clear existing content
                            document.body.innerHTML = '';
                            
                            // Create pages
                            pages.forEach(page => {
                                const pageDiv = document.createElement('div');
                                pageDiv.className = 'page';
                                pageDiv.innerHTML = page.content;
                                
                                const pageNum = document.createElement('div');
                                pageNum.className = 'page-number';
                                pageNum.textContent = page.number;
                                pageDiv.appendChild(pageNum);
                                
                                document.body.appendChild(pageDiv);
                            });
                        }
                        
                        window.addEventListener('load', function() {
                            setTimeout(renderPages, 100);
                        });
                      </script>
                    </head>
                    <body>
                      <div class="page">
                        $htmlContent
                      </div>
                    </body>
                    </html>
                """.trimIndent()

                webView.loadDataWithBaseURL(
                    null,
                    styledHtml,
                    "text/html",
                    "utf-8",
                    null
                )
            },
            modifier = Modifier
                .fillMaxSize()
                .padding(paddingValues)
        )
    }
}
