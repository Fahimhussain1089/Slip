package com.salty.payslip.utils


import android.content.Context
import com.salty.payslip.model.ProductItem
import org.apache.poi.ss.usermodel.*
import java.io.File
import java.io.FileOutputStream
import android.util.Log
import java.io.InputStream


class ExcelHelper(private val context: Context) {

    companion object {
        private const val TAG = "ExcelHelper"
    }

    fun updateInvoiceTemplate(
        clientName: String,
        productItems: List<ProductItem>,
        outputFilePath: String
    ): Boolean {
        var inputStream: InputStream? = null
        var workbook: Workbook? = null
        var outputStream: FileOutputStream? = null

        return try {
            Log.d(TAG, "Starting Excel export process...")
            Log.d(TAG, "Client: $clientName")
            Log.d(TAG, "Products: ${productItems.size}")
            Log.d(TAG, "Output path: $outputFilePath")

            // Check if template exists in assets
            val assetFiles = context.assets.list("")
            Log.d(TAG, "Assets files: ${assetFiles?.joinToString()}")

            // Read the existing template file from assets
            inputStream = context.assets.open("Invoice.xlsx")
            Log.d(TAG, "Template file opened successfully")

            workbook = WorkbookFactory.create(inputStream)
            val sheet = workbook.getSheetAt(0) // First sheet
            Log.d(TAG, "Workbook loaded with ${workbook.numberOfSheets} sheets")

            // Fill client name (M/s field) - Row 7, Column 1 (B7)
            setCellValue(sheet, 6, 1, clientName)
            Log.d(TAG, "Client name set: $clientName")

            // Fill product details starting from row 10
            var currentRow = 9 // Excel rows are 0-based, so row 9 = row 10 in Excel
            productItems.forEachIndexed { index, product ->
                Log.d(TAG, "Adding product ${index + 1}: ${product.name}")

                setCellValue(sheet, currentRow, 1, product.name)        // Description
                setCellValue(sheet, currentRow, 2, product.hsnCode)     // HSN Code
                setCellValue(sheet, currentRow, 3, product.quantity.toString()) // Qty
                setCellValue(sheet, currentRow, 4, product.price.toString())    // Rate

                // Calculate amount (Qty * Rate)
                val amount = product.quantity * product.price
                setCellValue(sheet, currentRow, 5, amount.toString())   // Amount
                Log.d(TAG, "Product added: ${product.name} - Qty: ${product.quantity}, Rate: ${product.price}, Amount: $amount")

                currentRow++
            }

            // Calculate subtotal
            val subtotal = productItems.sumOf { it.quantity * it.price }
            Log.d(TAG, "Subtotal calculated: $subtotal")

            // Create output directory if it doesn't exist
            val outputFile = File(outputFilePath)
            outputFile.parentFile?.mkdirs()
            Log.d(TAG, "Output directory prepared")

            // Save the workbook
            outputStream = FileOutputStream(outputFile)
            workbook.write(outputStream)
            Log.d(TAG, "Workbook written to output stream")

            // Close resources
            workbook.close()
            outputStream.close()
            inputStream.close()

            Log.d(TAG, "Excel export completed successfully")
            true

        } catch (e: Exception) {
            Log.e(TAG, "Error exporting to Excel", e)
            e.printStackTrace()

            // Close resources in case of error
            try {
                workbook?.close()
                outputStream?.close()
                inputStream?.close()
            } catch (closeEx: Exception) {
                Log.e(TAG, "Error closing resources", closeEx)
            }

            false
        }
    }

    private fun setCellValue(sheet: Sheet, rowNum: Int, colNum: Int, value: String) {
        try {
            var row = sheet.getRow(rowNum)
            if (row == null) {
                row = sheet.createRow(rowNum)
            }

            var cell = row.getCell(colNum)
            if (cell == null) {
                cell = row.createCell(colNum)
            }

            cell.setCellValue(value)
            Log.d(TAG, "Cell [$rowNum,$colNum] set to: $value")
        } catch (e: Exception) {
            Log.e(TAG, "Error setting cell value at [$rowNum,$colNum]", e)
        }
    }
}