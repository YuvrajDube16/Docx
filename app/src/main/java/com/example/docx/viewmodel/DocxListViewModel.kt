package com.example.docx.viewmodel

import android.app.Application
import android.content.ContentResolver
import android.net.Uri
import android.provider.MediaStore
import androidx.lifecycle.AndroidViewModel
import androidx.lifecycle.viewModelScope
import com.example.docx.data.DocxFileInfo
import com.example.docx.util.ProvidesLogging
import kotlinx.coroutines.Dispatchers
import kotlinx.coroutines.flow.MutableStateFlow
import kotlinx.coroutines.flow.StateFlow
import kotlinx.coroutines.launch
import kotlinx.coroutines.withContext
import java.io.File

class DocxListViewModel(application: Application) : AndroidViewModel(application), ProvidesLogging {
    private val _docxFiles = MutableStateFlow<List<DocxFileInfo>>(emptyList())
    val docxFiles: StateFlow<List<DocxFileInfo>> = _docxFiles

    init {
        viewModelScope.launch {
            scanForDocxFiles()
        }
    }

    fun scanForDocxFiles() {
        viewModelScope.launch {
            try {
                val files = withContext(Dispatchers.IO) {
                    val mediaStoreFiles = queryDocxFiles(getApplication<Application>().contentResolver)
                    logD("Found ${mediaStoreFiles.size} files through MediaStore")

                    if (mediaStoreFiles.isEmpty()) {
                        // Fallback to direct file system scan
                        val directFiles = scanFileSystemForDocx()
                        logD("Found ${directFiles.size} files through direct file system scan")
                        directFiles
                    } else {
                        mediaStoreFiles
                    }
                }
                _docxFiles.value = files
                logD("Total DOCX files found: ${files.size}")
            } catch (e: Exception) {
                logE("Error scanning for DOCX files", e)
            }
        }
    }

    private fun scanFileSystemForDocx(): List<DocxFileInfo> {
        val docxFiles = mutableListOf<DocxFileInfo>()
        val externalDirs = getApplication<Application>().getExternalFilesDirs(null)

        externalDirs.filterNotNull().forEach { dir ->
            // Get the root of external storage from the path
            var rootPath = dir.absolutePath
            val androidDirIndex = rootPath.indexOf("/Android")
            if (androidDirIndex > 0) {
                rootPath = rootPath.substring(0, androidDirIndex)
                scanDirectory(File(rootPath), docxFiles)
            }
        }

        return docxFiles
    }

    private fun scanDirectory(directory: File, results: MutableList<DocxFileInfo>) {
        try {
            directory.listFiles()?.forEach { file ->
                if (file.isDirectory) {
                    scanDirectory(file, results)
                } else if (file.name.endsWith(".docx", ignoreCase = true)) {
                    logD("Found DOCX file: ${file.name}")
                    results.add(
                        DocxFileInfo(
                            uri = Uri.fromFile(file),
                            name = file.name,
                            size = file.length(),
                            lastModified = file.lastModified(),
                            isFavorite = false
                        )
                    )
                }
            }
        } catch (e: Exception) {
            logE("Error scanning directory: ${directory.absolutePath}", e)
        }
    }

    private fun queryDocxFiles(contentResolver: ContentResolver): List<DocxFileInfo> {
        val docxFiles = mutableListOf<DocxFileInfo>()

        try {
            val collection = MediaStore.Files.getContentUri("external")

            val projection = arrayOf(
                MediaStore.Files.FileColumns._ID,
                MediaStore.Files.FileColumns.DISPLAY_NAME,
                MediaStore.Files.FileColumns.SIZE,
                MediaStore.Files.FileColumns.DATE_MODIFIED,
                MediaStore.Files.FileColumns.MIME_TYPE,
                MediaStore.Files.FileColumns.DATA
            )

            val selection = "(${MediaStore.Files.FileColumns.MIME_TYPE} = ? OR " +
                          "${MediaStore.Files.FileColumns.DISPLAY_NAME} LIKE ?) AND " +
                          "${MediaStore.Files.FileColumns.SIZE} > 0"

            val selectionArgs = arrayOf(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "%.docx"
            )

            val sortOrder = "${MediaStore.Files.FileColumns.DATE_MODIFIED} DESC"

            contentResolver.query(
                collection,
                projection,
                selection,
                selectionArgs,
                sortOrder
            )?.use { cursor ->
                logD("MediaStore query returned ${cursor.count} results")

                while (cursor.moveToNext()) {
                    val idColumn = cursor.getColumnIndexOrThrow(MediaStore.Files.FileColumns._ID)
                    val nameColumn = cursor.getColumnIndexOrThrow(MediaStore.Files.FileColumns.DISPLAY_NAME)
                    val sizeColumn = cursor.getColumnIndexOrThrow(MediaStore.Files.FileColumns.SIZE)
                    val dateColumn = cursor.getColumnIndexOrThrow(MediaStore.Files.FileColumns.DATE_MODIFIED)
                    val dataColumn = cursor.getColumnIndexOrThrow(MediaStore.Files.FileColumns.DATA)

                    val id = cursor.getLong(idColumn)
                    val name = cursor.getString(nameColumn)
                    val size = cursor.getLong(sizeColumn)
                    val lastModified = cursor.getLong(dateColumn) * 1000
                    val path = cursor.getString(dataColumn)

                    logD("Found DOCX: $name at $path")

                    val contentUri = Uri.withAppendedPath(collection, id.toString())
                    docxFiles.add(
                        DocxFileInfo(
                            uri = contentUri,
                            name = name,
                            size = size,
                            lastModified = lastModified,
                            isFavorite = false
                        )
                    )
                }
            }
        } catch (e: Exception) {
            logE("Error querying MediaStore", e)
        }

        return docxFiles
    }

    // Add this method to handle file list updates
    fun updateDocxFiles(files: List<DocxFileInfo>) {
        _docxFiles.value = files
    }

    fun refreshFiles() {
        scanForDocxFiles()
    }
}