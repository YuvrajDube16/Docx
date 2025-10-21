plugins {
    alias(libs.plugins.android.application)
    alias(libs.plugins.kotlin.android)
    alias(libs.plugins.kotlin.compose)
}

android {
    namespace = "com.example.docx"
    compileSdk = 36

    defaultConfig {
        applicationId = "com.example.docx"
        minSdk = 26 // Updated from 24 to 26 for Apache POI compatibility
        targetSdk = 36
        versionCode = 1
        versionName = "1.0"

        testInstrumentationRunner = "androidx.test.runner.AndroidJUnitRunner"
        vectorDrawables {
            useSupportLibrary = true
        }
    }

    buildTypes {
        release {
            isMinifyEnabled = false
            proguardFiles(
                getDefaultProguardFile("proguard-android-optimize.txt"),
                "proguard-rules.pro"
            )
        }
    }
    compileOptions {
        sourceCompatibility = JavaVersion.VERSION_11
        targetCompatibility = JavaVersion.VERSION_11
    }
    kotlinOptions {
        jvmTarget = "11"
    }
    buildFeatures {
        compose = true
    }
}

dependencies {

    implementation(libs.androidx.core.ktx)
    implementation(libs.androidx.lifecycle.runtime.ktx)
    implementation(libs.androidx.activity.compose)
    implementation(platform(libs.androidx.compose.bom))
    implementation(libs.androidx.compose.ui)
    implementation(libs.androidx.compose.ui.graphics)
    implementation(libs.androidx.compose.ui.tooling.preview)
    implementation(libs.androidx.compose.material3)
    implementation(libs.androidx.ui)
    implementation(libs.androidx.espresso.core)
    implementation(libs.androidx.ui.graphics)
    implementation(libs.androidx.compose.runtime.saveable)
    testImplementation(libs.junit)
    androidTestImplementation(libs.androidx.junit)
    androidTestImplementation(libs.androidx.espresso.core)
    androidTestImplementation(platform(libs.androidx.compose.bom))
    androidTestImplementation(libs.androidx.compose.ui.test.junit4)
    debugImplementation(libs.androidx.compose.ui.tooling)
    debugImplementation(libs.androidx.compose.ui.test.manifest)

//    // Apache POI dependencies
//    implementation("org.apache.poi:poi:5.2.5")
//    implementation("org.apache.poi:poi-ooxml:5.2.5")

    // Document picker
    implementation("androidx.activity:activity-ktx:1.8.2")
    implementation("androidx.documentfile:documentfile:1.0.1")

    // Navigation
    implementation("androidx.navigation:navigation-compose:2.7.6")

    // Gson for storing URIs in SharedPreferences
    implementation("com.google.code.gson:gson:2.10.1")

    // Google Accompanist for permissions
    implementation("com.google.accompanist:accompanist-permissions:0.32.0")

    // Jsoup for HTML parsing
    implementation("org.jsoup:jsoup:1.15.3")

    // Apache Commons Text for unescaping HTML
    implementation("org.apache.commons:commons-text:1.10.0")

    // Testing
    testImplementation("junit:junit:4.13.2")
    androidTestImplementation("androidx.test.ext:junit:1.1.5")
    androidTestImplementation("androidx.test.espresso:espresso-core:3.5.1")

    // Compose testing
    androidTestImplementation(platform("androidx.compose:compose-bom:2023.08.00"))
    androidTestImplementation("androidx.compose.ui:ui-test-junit4")

    // Debug dependencies
    debugImplementation("androidx.compose.ui:ui-tooling")
    debugImplementation("androidx.compose.ui:ui-test-manifest")

    // Coil for images in DOCX
    implementation("io.coil-kt:coil:2.4.0")

    // Coroutines
    implementation("org.jetbrains.kotlinx:kotlinx-coroutines-android:1.7.3")


    implementation("fr.opensagres.xdocreport:fr.opensagres.poi.xwpf.converter.xhtml:2.0.2"){
        exclude(group = "org.apache.poi", module = "ooxml-schemas")
    }
    val poiVersion = "5.2.3" // Use the latest stable version
    implementation("org.apache.poi:poi:$poiVersion")
    implementation("org.apache.poi:poi-ooxml:$poiVersion")

    // Apache POI dependencies for DOCX handling
    implementation("org.apache.poi:poi:5.2.3")
    implementation("org.apache.poi:poi-ooxml:5.2.3")

    // Required for Android compatibility
    implementation("org.apache.commons:commons-compress:1.24.0")
    implementation("org.apache.xmlbeans:xmlbeans:5.1.1")
}
