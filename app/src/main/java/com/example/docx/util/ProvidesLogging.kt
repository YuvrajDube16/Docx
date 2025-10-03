package com.example.docx.util

import android.util.Log

interface ProvidesLogging {
    val logTag: String
        get() = this.javaClass.simpleName

    fun logD(message: String) = Log.d(logTag, message)
    fun logE(message: String, throwable: Throwable? = null) = Log.e(logTag, message, throwable)
    fun logI(message: String) = Log.i(logTag, message)
}
