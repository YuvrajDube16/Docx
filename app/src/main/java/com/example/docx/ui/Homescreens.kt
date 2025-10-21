package com.example.docx.ui

import android.net.Uri
import androidx.compose.foundation.Image
import androidx.compose.foundation.clickable
import androidx.compose.foundation.layout.*
import androidx.compose.foundation.lazy.LazyColumn
import androidx.compose.foundation.lazy.items
import androidx.compose.material.icons.Icons
import androidx.compose.material.icons.automirrored.filled.ArrowBack
import androidx.compose.material.icons.filled.Add
import androidx.compose.material.icons.filled.MoreVert
import androidx.compose.material.icons.filled.Search
import androidx.compose.material.icons.filled.Star
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.graphics.Color
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.text.font.FontWeight
import androidx.compose.ui.text.style.TextOverflow
import androidx.compose.ui.unit.dp
import androidx.navigation.NavController
import com.example.docx.data.DocxFileInfo
import com.example.docx.navigation.Screens
import androidx.documentfile.provider.DocumentFile
import android.content.ContentUris
import android.provider.MediaStore
import android.os.Build
import android.widget.Toast
import androidx.compose.ui.res.painterResource
import androidx.compose.ui.tooling.preview.Preview
import androidx.lifecycle.viewmodel.compose.viewModel
import com.example.docx.R
import com.example.docx.util.Logger
import com.example.docx.viewmodel.DocxListViewModel
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext

@OptIn(ExperimentalMaterial3Api::class)
@Composable
fun HomeScreens(navController: NavController, viewModel: DocxListViewModel = viewModel()) {
    val context = LocalContext.current
    val scope = rememberCoroutineScope()
    val docxFiles by viewModel.docxFiles.collectAsState()
    var isLoading by remember { mutableStateOf(true) }
    var errorMessage by remember { mutableStateOf<String?>(null) }

    LaunchedEffect(Unit) {
        scope.launch {
            try {
                val files = withContext(Dispatchers.IO) {
                    val collection = if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                        MediaStore.Files.getContentUri(MediaStore.VOLUME_EXTERNAL)
                    } else {
                        MediaStore.Files.getContentUri("external")
                    }

                    val projection = arrayOf(
                        MediaStore.Files.FileColumns._ID,
                        MediaStore.Files.FileColumns.DISPLAY_NAME,
                        MediaStore.Files.FileColumns.SIZE,
                        MediaStore.Files.FileColumns.DATE_MODIFIED,
                        MediaStore.Files.FileColumns.MIME_TYPE
                    )

                    val selection = "${MediaStore.Files.FileColumns.MIME_TYPE} = ? OR " +
                            "${MediaStore.Files.FileColumns.DISPLAY_NAME} LIKE ?"
                    val selectionArgs = arrayOf(
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        "%.docx"
                    )

                    val docxList = mutableListOf<DocxFileInfo>()

                    context.contentResolver.query(
                        collection,
                        projection,
                        selection,
                        selectionArgs,
                        "${MediaStore.Files.FileColumns.DATE_MODIFIED} DESC"
                    )?.use { cursor ->
                        val idColumn = cursor.getColumnIndexOrThrow(MediaStore.Files.FileColumns._ID)
                        val nameColumn = cursor.getColumnIndexOrThrow(MediaStore.Files.FileColumns.DISPLAY_NAME)
                        val sizeColumn = cursor.getColumnIndexOrThrow(MediaStore.Files.FileColumns.SIZE)
                        val dateColumn = cursor.getColumnIndexOrThrow(MediaStore.Files.FileColumns.DATE_MODIFIED)

                        while (cursor.moveToNext()) {
                            val id = cursor.getLong(idColumn)
                            val name = cursor.getString(nameColumn)
                            val size = cursor.getLong(sizeColumn)
                            val date = cursor.getLong(dateColumn)

                            val contentUri = ContentUris.withAppendedId(collection, id)
                            val documentFile = DocumentFile.fromSingleUri(context, contentUri)

                            if (documentFile?.exists() == true) {
                                docxList.add(
                                    DocxFileInfo(
                                        uri = contentUri,
                                        name = name,
                                        size = size,
                                        lastModified = date * 1000,
                                        isFavorite = false
                                    )
                                )
                            }
                        }
                    }
                    docxList
                }
                viewModel.updateDocxFiles(files)
            } catch (e: Exception) {
                errorMessage = "Unable to access documents. Please check permissions."
            } finally {
                isLoading = false
            }
        }
    }

    // Show error message if any
    errorMessage?.let { message ->
        LaunchedEffect(message) {
            Toast.makeText(context, message, Toast.LENGTH_LONG).show()
            errorMessage = null
        }
    }

    Scaffold(
        topBar = {
            TopAppBar(
                title = {
                    Text("Word Files", fontWeight = FontWeight.Bold)
                },
                navigationIcon = {
                    IconButton(onClick = { navController.popBackStack() }) {
                        Icon(Icons.AutoMirrored.Filled.ArrowBack, contentDescription = "Back")
                    }
                },
                actions = {
                    // AZ Sorting Icon
                    IconButton(onClick = { /* Handle sort */ }) {
                        Text("AZ", style = MaterialTheme.typography.titleMedium, fontWeight = FontWeight.Bold)
                    }
                    // Search/Filter Icon
                    IconButton(onClick = { /* Handle search */ }) {
                        Icon(Icons.Default.Search, contentDescription = "Search")
                    }
                },
                colors = TopAppBarDefaults.topAppBarColors(
                    containerColor = MaterialTheme.colorScheme.surfaceColorAtElevation(3.dp) // Subtle elevation
                )
            )
        },
        floatingActionButton = {
            FloatingActionButton(onClick = { navController.navigate(Screens.CreateNewDocx.route) }) {
                Icon(Icons.Filled.Add, contentDescription = "Create new document")
            }
        }
    ) { paddingValues ->
        if (docxFiles.isEmpty()) {
            Box(
                modifier = Modifier
                    .fillMaxSize()
                    .padding(paddingValues),
                contentAlignment = Alignment.Center
            ) {
                Column(
                    horizontalAlignment = Alignment.CenterHorizontally,
                    modifier = Modifier.padding(16.dp)
                ) {
                    Text("No DOCX files found.", style = MaterialTheme.typography.titleMedium)
                    Spacer(modifier = Modifier.height(8.dp))
                    Text("Please ensure you've granted access to your main 'Documents' directory in the initial setup.",
                        style = MaterialTheme.typography.bodyMedium,
                        color = MaterialTheme.colorScheme.onSurfaceVariant
                    )
                }
            }
        } else {
            LazyColumn(
                modifier = Modifier
                    .fillMaxSize()
                    .padding(paddingValues)
                    .padding(horizontal = 8.dp, vertical = 4.dp), // Adjust padding for cards
                verticalArrangement = Arrangement.spacedBy(8.dp) // Spacing between cards
            ) {
                items(docxFiles) { fileInfo ->
                    DocxFileCard(
                        fileInfo = fileInfo,
                        onClick = {
                            try {
                                Logger.d("Attempting to open document: ${fileInfo.name}")
                                val encodedUri = Uri.encode(fileInfo.uri.toString())
                                navController.navigate("view/$encodedUri")
                            } catch (e: Exception) {
                                Logger.e("Error navigating to document", e)
                                Toast.makeText(
                                    context,
                                    "Unable to open document: ${e.message}",
                                    Toast.LENGTH_SHORT
                                ).show()
                            }
                        }
                    )
                }
            }
        }
    }
}

@Composable
fun DocxFileCard(fileInfo: DocxFileInfo, onClick: () -> Unit) {
    Card(
        modifier = Modifier
            .fillMaxWidth()
            .clickable(onClick = onClick),
        elevation = CardDefaults.cardElevation(defaultElevation = 2.dp),
        colors = CardDefaults.cardColors(containerColor = MaterialTheme.colorScheme.surface)
    ) {
        Row(
            modifier = Modifier
                .fillMaxWidth()
                .padding(16.dp),
            verticalAlignment = Alignment.CenterVertically
        ) {
            // DOCX Icon
            Image(
                painter = painterResource(id = R.drawable.docx_svg), // You'll need to add this drawable
                contentDescription = "DOCX File",
                modifier = Modifier.size(48.dp) // Adjust size as needed
            )
            Spacer(modifier = Modifier.width(16.dp))

            // File Name and Metadata
            Column(modifier = Modifier.weight(1f)) {
                Text(
                    text = fileInfo.name,
                    style = MaterialTheme.typography.titleMedium,
                    fontWeight = FontWeight.Bold,
                    maxLines = 1,
                    overflow = TextOverflow.Ellipsis
                )
                Spacer(modifier = Modifier.height(4.dp))
                Row(verticalAlignment = Alignment.CenterVertically) {
                    Text(
                        text = fileInfo.getFormattedLastModified(),
                        style = MaterialTheme.typography.bodySmall,
                        color = MaterialTheme.colorScheme.onSurfaceVariant
                    )
                    Spacer(modifier = Modifier.width(8.dp))
                    Text(
                        text = fileInfo.getFormattedSize(),
                        style = MaterialTheme.typography.bodySmall,
                        color = MaterialTheme.colorScheme.onSurfaceVariant
                    )
                }
            }

            // Right icons (Star and MoreVert)
            Row(verticalAlignment = Alignment.CenterVertically) {
                IconButton(onClick = { /* Handle favorite toggle */ }) {
                    // Use a filled star for favorited, border for not
                    Icon(
                        imageVector = Icons.Filled.Star
//                            if (fileInfo.isFavorite) Icons.Filled.Star
//                        else Icons.Default.StarBorder
                        ,
                        contentDescription = "Favorite",
                        tint = if (fileInfo.isFavorite) Color(0xFFFFC107) else MaterialTheme.colorScheme.onSurfaceVariant // Yellow for favorite
                    )
                }
                IconButton(onClick = { /* Handle more options */ }) {
                    Icon(Icons.Default.MoreVert, contentDescription = "More options")
                }
            }
        }
    }
}