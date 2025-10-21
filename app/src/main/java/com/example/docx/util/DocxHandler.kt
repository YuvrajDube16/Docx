package com.example.docx.util

import android.content.Context
import android.net.Uri
import org.apache.poi.xwpf.usermodel.XWPFDocument
import org.apache.poi.xwpf.usermodel.XWPFParagraph
import kotlin.math.absoluteValue

class DocxHandler {
    companion object {
        // Minimum allowed line height
        private const val MIN_LINE_HEIGHT = 0.1

        /**
         * Reads a DOCX file from a Uri and returns an XWPFDocument
         */
        fun readDocx(context: Context, uri: Uri): XWPFDocument {
            return context.contentResolver.openInputStream(uri)?.use { inputStream ->
                XWPFDocument(inputStream)
            } ?: throw IllegalStateException("Could not open document")
        }

        /**
         * Fixes negative line height values in the document
         */
        private fun fixLineHeights(document: XWPFDocument) {
            document.paragraphs.forEach { paragraph ->
                val spacing = paragraph.spacingBetween
                if (spacing < 0) {
                    // If spacing is negative, set it to its absolute value or minimum allowed
                    paragraph.spacingBetween = spacing.absoluteValue.coerceAtLeast(MIN_LINE_HEIGHT)
                }
            }
        }

        /**
         * Saves the XWPFDocument to a new file at the specified Uri
         * This method ensures the file is saved in a format compatible with other DOCX readers
         */
        fun saveDocx(context: Context, document: XWPFDocument, outputUri: Uri) {
            // Fix any negative line heights before saving
            fixLineHeights(document)

            context.contentResolver.openOutputStream(outputUri)?.use { outputStream ->
                // Write the document using POI's built-in serialization
                // This maintains the ZIP structure that other readers expect
                document.write(outputStream)
                outputStream.flush()
            } ?: throw IllegalStateException("Could not save document")
        }

        /**
         * Edits the content of a DOCX file
         * @param sourceUri The Uri of the source DOCX file
         * @param destinationUri The Uri where the edited file should be saved
         * @param editOperation A lambda that performs the desired edits on the XWPFDocument
         */
        fun editDocx(
            context: Context,
            sourceUri: Uri,
            destinationUri: Uri,
            editOperation: (XWPFDocument) -> Unit
        ) {
            val document = readDocx(context, sourceUri)
            try {
                // Perform the edit operation
                editOperation(document)

                // Fix any negative line heights that might have been introduced
                fixLineHeights(document)

                // Save the modified document
                saveDocx(context, document, destinationUri)
            } finally {
                // Clean up
                document.close()
            }
        }

        /**
         * Reads and extracts text content from a DOCX file
         * @return A list of paragraphs from the document
         */
        fun readDocxContent(context: Context, uri: Uri): List<String> {
            return readDocx(context, uri).use { document ->
                document.paragraphs.map { paragraph ->
                    paragraph.text
                }
            }
        }

        /**
         * Verifies if the file is a readable DOCX file
         * @return true if the file can be read as a DOCX
         */
        fun isReadableDocx(context: Context, uri: Uri): Boolean {
            return try {
                readDocx(context, uri).use { document ->
                    // Try to access paragraphs to verify document structure
                    document.paragraphs
                    true
                }
            } catch (e: Exception) {
                e.printStackTrace()
                false
            }
        }
    }
}
