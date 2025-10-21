package com.example.docx.repository

import android.content.Context
import android.net.Uri
import androidx.core.content.edit

class DocxRepository(private val context: Context) {
    private val prefs = context.getSharedPreferences("docx_prefs", Context.MODE_PRIVATE)
    private val favoritePrefs = context.getSharedPreferences("docx_favorites", Context.MODE_PRIVATE)

    fun getStoredDocxUris(): List<Uri> {
        return prefs.getStringSet("docx_uris", emptySet())
            ?.mapNotNull { Uri.parse(it) }
            ?: emptyList()
    }

    fun addDocxUri(uri: Uri) {
        val currentUris = getStoredDocxUris().toMutableSet()
        currentUris.add(uri)
        prefs.edit {
            putStringSet("docx_uris", currentUris.map { it.toString() }.toSet())
        }
    }

    fun getDocumentsDirectoryUri(): Uri? {
        val uriString = prefs.getString("documents_dir_uri", null)
        return uriString?.let { Uri.parse(it) }
    }

    fun saveDocumentsDirectoryUri(uri: Uri) {
        prefs.edit {
            putString("documents_dir_uri", uri.toString())
        }
    }

    fun isFileFavorite(uri: Uri): Boolean {
        return favoritePrefs.getBoolean(uri.toString(), false)
    }

    fun toggleFileFavorite(uri: Uri) {
        val currentState = isFileFavorite(uri)
        favoritePrefs.edit {
            putBoolean(uri.toString(), !currentState)
        }
    }
}