plugins {
    alias(libs.plugins.android.application)
    alias(libs.plugins.kotlin.android)
    id("kotlin-parcelize")
}

android {
    namespace = "com.salty.payslip"
    compileSdk = 36

    defaultConfig {
        applicationId = "com.salty.payslip"
        minSdk = 24
        targetSdk = 36
        versionCode = 1
        versionName = "1.0"

        testInstrumentationRunner = "androidx.test.runner.AndroidJUnitRunner"
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
        viewBinding = true
    }

    packagingOptions {
        resources.excludes.apply {
            add("META-INF/DEPENDENCIES")
            add("META-INF/LICENSE")
            add("META-INF/LICENSE.txt")
            add("META-INF/NOTICE")
            add("META-INF/NOTICE.txt")
            add("META-INF/INDEX.LIST")
            add("META-INF/io.netty.versions.properties")
        }
    }

}

dependencies {

    implementation(libs.androidx.core.ktx)
    implementation(libs.androidx.appcompat)
    implementation(libs.material)
    implementation(libs.androidx.constraintlayout)
    implementation(libs.androidx.navigation.fragment.ktx)
    implementation(libs.androidx.navigation.ui.ktx)
    implementation(libs.filament.android)
    testImplementation(libs.junit)
    androidTestImplementation(libs.androidx.junit)
    androidTestImplementation(libs.androidx.espresso.core)


    // For reading Excel files
//    implementation("org.apache.poi:poi:5.2.3")
//    implementation("org.apache.poi:poi-ooxml:5.2.3")
//
//    // For file picker - updated version
//    implementation("com.github.dhaval2404:imagepicker:2.1")
//
//    // For permissions
//    implementation("com.karumi:dexter:6.2.3")
//
//    // For displaying data in table
//    implementation("androidx.recyclerview:recyclerview:1.3.2")
//
//    // Add these if not already present
//    implementation("androidx.lifecycle:lifecycle-viewmodel-ktx:2.7.0")
//    implementation("androidx.activity:activity-ktx:1.8.2")


//    // Alternative: Use a simpler Excel reader library
//    implementation("com.github.andruhon:Android5xLS:1.0")
//
//    // Or use CSV reader only (simpler approach)
//    implementation("com.github.doyaaaaaken:kotlin-csv-jvm:1.9.2")
//
//    // Keep the rest of your dependencies
//    implementation("com.github.dhaval2404:imagepicker:2.1")
//    implementation("com.karumi:dexter:6.2.3")
//    implementation("androidx.recyclerview:recyclerview:1.3.2")

    ///*************new approach

    // For displaying data in table
    implementation("androidx.recyclerview:recyclerview:1.3.2")

    // Add these lifecycle components
    implementation("androidx.lifecycle:lifecycle-viewmodel-ktx:2.7.0")
    implementation("androidx.activity:activity-ktx:1.8.2")
    implementation("androidx.fragment:fragment-ktx:1.6.2")

    // Add Navigation Safe Args
    implementation("androidx.navigation:navigation-fragment-ktx:2.7.7")
    implementation("androidx.navigation:navigation-ui-ktx:2.7.7")

}