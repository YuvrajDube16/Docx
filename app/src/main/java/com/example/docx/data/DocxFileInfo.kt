package com.example.docx.data

import android.net.Uri
import java.text.SimpleDateFormat
import java.util.*
import kotlin.math.log10
import kotlin.math.pow

data class DocxFileInfo(
    val uri: Uri,
    val name: String,
    val lastModified: Long = 0,
    val size: Long = 0,
    val isFavorite: Boolean = false
) {
    fun getFormattedLastModified(): String {
        if (lastModified == 0L) return "Unknown date"
        val sdf = SimpleDateFormat("MMM d, yyyy", Locale.getDefault())
        return sdf.format(Date(lastModified))
    }

    fun getFormattedSize(): String {
        if (size <= 0) return "0 B"
        val units = arrayOf("B", "KB", "MB", "GB", "TB")
        val digitGroups = (log10(size.toDouble()) / log10(1024.0)).toInt()
        return String.format("%.1f %s", size / 1024.0.pow(digitGroups.toDouble()), units[digitGroups])
    }
}
