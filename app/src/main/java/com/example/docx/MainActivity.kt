package com.example.docx

import android.Manifest
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.provider.Settings
import android.util.Base64
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.layout.*
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.style.TextAlign
import androidx.compose.ui.unit.dp
import androidx.lifecycle.Lifecycle
import androidx.lifecycle.LifecycleEventObserver
import androidx.lifecycle.viewmodel.compose.viewModel
import androidx.navigation.NavType
import androidx.navigation.compose.NavHost
import androidx.navigation.compose.composable
import androidx.navigation.compose.rememberNavController
import androidx.navigation.navArgument
import com.example.docx.navigation.Screens
import com.example.docx.ui.CreateNewDocxScreen
import com.example.docx.ui.EditDocxScreen
import com.example.docx.ui.HomeScreens
import com.example.docx.ui.ViewDocxScreen
import com.example.docx.ui.theme.DocxTheme
import com.example.docx.util.Logger
import com.example.docx.viewmodel.DocxListViewModel
import com.google.accompanist.permissions.ExperimentalPermissionsApi
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.withContext
import org.apache.poi.common.usermodel.PictureType
import org.apache.poi.util.Units
import org.apache.poi.xwpf.usermodel.*
import org.openxmlformats.schemas.officeDocument.x2006.sharedTypes.STVerticalAlignRun as OfficeSTVerticalAlignRun
import java.io.IOException
import java.net.URLEncoder // Added for URL encoding
import java.nio.charset.StandardCharsets // Added for Charset

class MainActivity : ComponentActivity() {
    private var permissionCallback: (() -> Unit)? = null

    private val requestPermissionLauncher = registerForActivityResult(
        ActivityResultContracts.RequestPermission()
    ) { isGranted: Boolean ->
        if (isGranted) {
            permissionCallback?.invoke()
        }
    }

    @OptIn(ExperimentalPermissionsApi::class)
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)

        setContent {
            DocxTheme {
                Surface(
                    modifier = Modifier.fillMaxSize(),
                    color = MaterialTheme.colorScheme.background
                ) {
                    val navController = rememberNavController()
                    val viewModel: DocxListViewModel = viewModel()
                    var hasPermission by remember { mutableStateOf(checkStoragePermission()) }
                    val context = LocalContext.current

                    DisposableEffect(Unit) {
                        val observer = LifecycleEventObserver { _, event ->
                            if (event == Lifecycle.Event.ON_RESUME) {
                                val currentPermissionState = checkStoragePermission()
                                if (currentPermissionState && !hasPermission) {
                                    hasPermission = true
                                    viewModel.scanForDocxFiles()
                                }
                            }
                        }
                        lifecycle.addObserver(observer)
                        onDispose {
                            lifecycle.removeObserver(observer)
                        }
                    }

                    NavHost(navController = navController, startDestination = if (hasPermission) "home" else "permission") {
                        composable("permission") {
                            PermissionScreen(onRequestPermission = {
                                requestStoragePermission {
                                    hasPermission = true
                                    viewModel.scanForDocxFiles()
                                    navController.navigate("home") {
                                        popUpTo("permission") { inclusive = true }
                                    }
                                }
                            })
                        }
                        composable("home") {
                            HomeScreens(navController = navController, viewModel = viewModel)
                        }
                        composable(
                            route = "view/{fileUri}",
                            arguments = listOf(navArgument("fileUri") { type = NavType.StringType })
                        ) { backStackEntry ->
                            val fileUriString = backStackEntry.arguments?.getString("fileUri")
                            if (fileUriString != null) {
                                val docUri = Uri.parse(fileUriString) // The URI string from argument is already decoded by Nav
                                var htmlContent by remember { mutableStateOf("<p></p>") }
                                var isLoadingHtml by remember { mutableStateOf(true) }
                                var errorLoadingHtml by remember { mutableStateOf<String?>(null) }

                                LaunchedEffect(docUri) {
                                    isLoadingHtml = true
                                    errorLoadingHtml = null
                                    withContext(Dispatchers.IO) {
                                        try {
                                            context.contentResolver.openInputStream(docUri)?.use { inputStream ->
                                                val document = XWPFDocument(inputStream)
                                                val generatedHtml = convertDocxToHtmlEnhanced(document) // Helper function call
                                                withContext(Dispatchers.Main) {
                                                    htmlContent = generatedHtml
                                                    isLoadingHtml = false
                                                }
                                                document.close()
                                            } ?: throw IOException("Failed to open input stream for $docUri")
                                        } catch (e: Exception) {
                                            Logger.e("Error loading DOCX for view: ${e.localizedMessage}", e)
                                            withContext(Dispatchers.Main) {
                                                htmlContent = "<p>Error loading document: ${e.localizedMessage}</p>"
                                                errorLoadingHtml = "Failed to load document: ${e.localizedMessage}"
                                                isLoadingHtml = false
                                            }
                                        }
                                    }
                                }

                                if (isLoadingHtml) {
                                    Box(modifier = Modifier.fillMaxSize(), contentAlignment = Alignment.Center) {
                                        CircularProgressIndicator()
                                    }
                                } else if (errorLoadingHtml != null) {
                                    Box(modifier = Modifier.fillMaxSize().padding(16.dp), contentAlignment = Alignment.Center) {
                                        Text(errorLoadingHtml!!, color = MaterialTheme.colorScheme.error, textAlign = TextAlign.Center)
                                    }
                                } else {
                                    ViewDocxScreen(
                                        htmlContent = htmlContent,
                                        fileUriString = fileUriString, // Pass the original, non-encoded URI string to ViewDocxScreen
                                        onNavigateBack = { navController.navigateUp() },
                                        onNavigateToEdit = { uriToEdit -> // This is the original fileUriString
                                            val encodedUri = URLEncoder.encode(uriToEdit, StandardCharsets.UTF_8.toString())
                                            navController.navigate("edit/$encodedUri") 
                                        }
                                    )
                                }
                            } else {
                                Text("Error: File URI not provided.")
                            }
                        }
                        composable(
                            route = "edit/{fileUri}",
                            arguments = listOf(navArgument("fileUri") { type = NavType.StringType })
                        ) { backStackEntry ->
                            // The fileUri argument is automatically URL-decoded by Jetpack Navigation
                            backStackEntry.arguments?.getString("fileUri")?.let { uriString -> 
                                EditDocxScreen(
                                    navController = navController,
                                    fileUriString = uriString
                                )
                            }
                        }
                        composable(
                            route = Screens.CreateNewDocx.route,
                        ) {
                            CreateNewDocxScreen(navController)
                        }
                    }
                }
            }
        }
    }

    private fun checkStoragePermission(): Boolean {
        return if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R) {
            Environment.isExternalStorageManager()
        } else {
            checkSelfPermission(Manifest.permission.READ_EXTERNAL_STORAGE) == PackageManager.PERMISSION_GRANTED
        }
    }

    private fun requestStoragePermission(onGranted: () -> Unit) {
        permissionCallback = onGranted
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.R) {
            try {
                val intent = Intent(Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION).apply {
                    data = Uri.parse("package:$packageName")
                }
                startActivity(intent)
            } catch (e: Exception) {
                Logger.e("Error requesting MANAGE_APP_ALL_FILES_ACCESS_PERMISSION: ${e.message}", e)
                // Fallback or inform user if the intent fails
                val intent = Intent(Settings.ACTION_MANAGE_ALL_FILES_ACCESS_PERMISSION)
                startActivity(intent)
            }
        } else {
            requestPermissionLauncher.launch(Manifest.permission.READ_EXTERNAL_STORAGE)
        }
    }

    companion object {
        private const val PERMISSION_REQUEST_CODE = 100
    }
}

@Composable
private fun PermissionScreen(onRequestPermission: () -> Unit) {
    Column(
        modifier = Modifier
            .fillMaxSize()
            .padding(24.dp),
        horizontalAlignment = Alignment.CenterHorizontally,
        verticalArrangement = Arrangement.Center
    ) {
        Text(
            text = "Storage Access Required",
            style = MaterialTheme.typography.headlineSmall
        )
        Spacer(modifier = Modifier.height(16.dp))
        Text(
            text = "This app needs access to storage to find and display DOCX files on your device.",
            textAlign = TextAlign.Center,
            style = MaterialTheme.typography.bodyLarge
        )
        Spacer(modifier = Modifier.height(24.dp))
        Button(onClick = onRequestPermission) {
            Text("Grant Permission")
        }
    }
}

// Copied DOCX -> HTML Conversion Helper Functions (from EditDocxScreen.kt)
fun convertDocxToHtmlEnhanced(document: XWPFDocument): String {
    val html = StringBuilder()
    html.append("<!DOCTYPE html><html><head>")
    html.append("<meta charset=\"UTF-8\">")
    html.append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">")
    html.append("<style>")
    html.append("body { font-family: Arial, sans-serif; font-size: 12pt; margin: 12px; word-wrap: break-word; background: #fff; }")
    html.append("p { margin-bottom: 0.5em; }")
    html.append("table { border-collapse: collapse; width: 100%; margin-bottom: 1em; table-layout: fixed; }")
    html.append("td, th { border: 1px solid #BFBFBF; padding: 6px 8px; vertical-align: top; text-align: left; word-wrap: break-word; font-size: 10pt; }")
    html.append("img { max-width: 100%; height: auto; display: block; margin: 0.2em 0; }")
    html.append("</style>")
    html.append("</head><body>")

    for (bodyElement in document.bodyElements) {
        when (bodyElement) {
            is XWPFParagraph -> html.append(convertParagraphToHtmlExact(bodyElement))
            is XWPFTable -> html.append(convertTableToHtmlExact(bodyElement))
            // Note: XWPFSDT (Structured Document Tag) elements might also exist and might need handling
        }
    }

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
        // Preserve empty paragraphs if they have specific spacing that might be intentional
        // Or use &nbsp; if they are truly meant to be just a visual space with no semantic meaning.
        // For now, let's keep it simple and allow them to render as potentially collapsed empty <p> tags.
        // If they need to occupy space, CSS like min-height or adding &nbsp; would be needed.
        pTag.append("&nbsp;") // To ensure paragraph takes up some space
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
        // Note: For more complex underline patterns, specific CSS might be needed.
    }
    var text = run.text() ?: ""
    // HTML escape the text
    text = text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\"", "&quot;")
        .replace("\'", "&#39;")
        .replace("\t", "&nbsp;&nbsp;&nbsp;&nbsp;") // Convert tabs to spaces
        .replace("\n", "<br>") // Convert newlines to <br>

    var html = if (style.isNotEmpty()) "<span style='${style}'>$text</span>" else text

    // Handle vertical alignment (subscript/superscript)
    when (run.getVerticalAlignment()) {
        OfficeSTVerticalAlignRun.SUBSCRIPT -> html = "<sub>$html</sub>"
        OfficeSTVerticalAlignRun.SUPERSCRIPT -> html = "<sup>$html</sup>"
        else -> {}
    }

    // Handle embedded pictures
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
                else -> "png" // Default to PNG
            }
            val base64 = Base64.encodeToString(picData.data, Base64.DEFAULT)
            
            var imgWidth = "auto"
            var imgHeight = "auto"
            // Try to get dimensions from a more reliable source if available in XWPFRun or XWPFPicture
            // For now, using a default or potentially from CTDrawing if accessible directly
            // pic.ctPicture.spPr.xfrm.ext.cx and cy are in EMUs
            val cx = pic.ctPicture?.spPr?.xfrm?.ext?.cx
            val cy = pic.ctPicture?.spPr?.xfrm?.ext?.cy
            if (cx != null && cx > 0) imgWidth = "${cx / Units.EMU_PER_PIXEL}px"
            if (cy != null && cy > 0) imgHeight = "${cy / Units.EMU_PER_PIXEL}px"

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
            // Basic cell properties could be handled here (width, background color, etc.)
            // val cellWidth = cell.width // Example
            // html.append("<td style='width:${cellWidth}pct;'>")
            html.append("<td>")
            cell.bodyElements.forEach { element ->
                when (element) {
                    is XWPFParagraph -> html.append(convertParagraphToHtmlExact(element))
                    is XWPFTable -> html.append(convertTableToHtmlExact(element)) // Nested tables
                }
            }
            html.append("</td>")
        }
        html.append("</tr>")
    }
    html.append("</table>")
    return html.toString()
}
