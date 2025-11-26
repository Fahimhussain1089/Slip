package com.salty.payslip.model

import android.os.Parcelable
import kotlinx.parcelize.Parcelize

// Product.kt

//data class Client(
//    val name: String
//)
//
//data class ProductItem(
//    val name: String,
//    val description: String = "",
//    val price: Double = 0.0
//)
//
//// Main data class to hold both lists
//data class FileData(
//    val clients: List<Client> = emptyList(),
//    val products: List<ProductItem> = emptyList()
//)

//****** different file upload **********
@Parcelize
data class Client(
    val name: String,
    val email: String = "",
    val phone: String = ""
) : Parcelable

@Parcelize
data class ProductItem(
    val name: String,
    val description: String = "",
    val price: Double = 0.0,
    val quantity: Int = 0,
    val hsnCode: String = ""
) : Parcelable