package com.example.docx

import android.Manifest
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.provider.Settings
import androidx.activity.ComponentActivity
import androidx.activity.compose.setContent
import androidx.activity.result.contract.ActivityResultContracts
import androidx.compose.foundation.layout.Arrangement
import androidx.compose.foundation.layout.Box
import androidx.compose.foundation.layout.Column
import androidx.compose.foundation.layout.Spacer
import androidx.compose.foundation.layout.fillMaxSize
import androidx.compose.foundation.layout.height
import androidx.compose.foundation.layout.padding
import androidx.compose.material3.*
import androidx.compose.runtime.*
import androidx.compose.ui.Alignment
import androidx.compose.ui.Modifier
import androidx.compose.ui.platform.LocalContext
import androidx.compose.ui.platform.LocalLifecycleOwner
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
import java.net.URLEncoder
import java.nio.charset.StandardCharsets

class MainActivity : ComponentActivity() {
    private val requestPermissionLauncher = registerForActivityResult(ActivityResultContracts.RequestPermission()) { isGranted: Boolean ->
        if (isGranted) {
            permissionCallback?.invoke()
        }
    }

    private var permissionCallback: (() -> Unit)? = null

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContent {
            DocxTheme {
                Surface(modifier = Modifier.fillMaxSize(), color = MaterialTheme.colorScheme.background) {
                    val navController = rememberNavController()
                    val context = LocalContext.current
                    val viewModel: DocxListViewModel = viewModel()
                    var hasPermission by remember { mutableStateOf(checkStoragePermission()) }

                    val lifecycle = LocalLifecycleOwner.current.lifecycle
                    DisposableEffect(lifecycle) {
                        val observer = LifecycleEventObserver { _, event ->
                            if (event == Lifecycle.Event.ON_RESUME) {
                                val permissionGranted = checkStoragePermission()
                                if (permissionGranted != hasPermission) {
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
                                ViewDocxScreen(
                                    fileUriString = fileUriString,
                                    onNavigateBack = { navController.navigateUp() },
                                    onNavigateToEdit = { uriToEdit ->
                                        val encodedUri = URLEncoder.encode(uriToEdit, StandardCharsets.UTF_8.toString())
                                        navController.navigate("edit/$encodedUri")
                                    }
                                )
                            } else {
                                Text("Error: File URI not provided.")
                            }
                        }
                        composable(
                            route = "edit/{fileUri}",
                            arguments = listOf(navArgument("fileUri") { type = NavType.StringType })
                        ) { backStackEntry ->
                            backStackEntry.arguments?.getString("fileUri")?.let { uriString ->
                                EditDocxScreen(
                                    fileUriString = uriString,
                                    onNavigateBack = { navController.navigateUp() }
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
                val intent = Intent(Settings.ACTION_MANAGE_ALL_FILES_ACCESS_PERMISSION)
                startActivity(intent)
            }
        } else {
            requestPermissionLauncher.launch(Manifest.permission.READ_EXTERNAL_STORAGE)
        }
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
