package com.example.docx.navigation

sealed class Screens(val route: String) {
    object Home : Screens("home")
    object ViewDocx : Screens("view-docx") // Existing screen
    object EditDocx : Screens("edit-docx") // Existing screen
    object CreateNewDocx : Screens("CreateNewDocxScreen") // New screen
}
