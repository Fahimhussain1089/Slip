package com.salty.payslip.Fragment

import android.Manifest
import android.content.pm.PackageManager
import android.os.Bundle
import android.util.Log
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.FrameLayout
import android.widget.ImageButton
import android.widget.Toast
import androidx.core.content.ContextCompat
import androidx.fragment.app.Fragment
import androidx.navigation.fragment.findNavController
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import com.salty.payslip.Adapter.DropdownAdapter
import com.salty.payslip.R
import com.salty.payslip.databinding.FragmentSecondBinding
import com.salty.payslip.model.Client
import com.salty.payslip.model.ProductItem
import java.io.File
import android.content.ContentValues
import android.os.Build
import android.os.Environment
import android.provider.MediaStore
import java.io.OutputStream
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

//========by cloud ai
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.common.usermodel.HyperlinkType
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import java.io.FileOutputStream
import android.app.AlertDialog
import android.widget.Switch
import android.widget.TextView
import com.google.android.material.textfield.TextInputEditText


class SecondFragment : Fragment() {

    private var _binding: FragmentSecondBinding? = null
    private val binding get() = _binding!!

    private val clientList = mutableListOf<Client>()
    private val productList = mutableListOf<ProductItem>()

    private var isClientDropdownOpen = false
    private var isProductDropdownOpen = false

    private lateinit var clientDropdownView: View
    private lateinit var productDropdownView: View

    private var selectedClient: Client? = null
    private var selectedProduct: ProductItem? = null

    // Invoice Data Variables
    private var invoiceDate: String = SimpleDateFormat("dd/MM/yyyy", Locale.getDefault()).format(Date())
    private var cgstRate: Double = 9.0
    private var sgstRate: Double = 9.0
    private var igstRate: Double = 0.0
    private var isReverseCharge: Boolean = false

    companion object {
        private const val STORAGE_PERMISSION_CODE = 1001
        private const val TAG = "SecondFragment"
    }

    override fun onCreateView(
        inflater: LayoutInflater, container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View {
        _binding = FragmentSecondBinding.inflate(inflater, container, false)
        return binding.root
    }

    override fun onViewCreated(view: View, savedInstanceState: Bundle?) {
        super.onViewCreated(view, savedInstanceState)

        // Get data from arguments (Bundle)
        arguments?.let { bundle ->
            bundle.getParcelableArrayList<Client>("clientList")?.let {
                clientList.clear()
                clientList.addAll(it)
            }
            bundle.getParcelableArrayList<ProductItem>("productList")?.let {
                productList.clear()
                productList.addAll(it)
            }
        }

        setupDropdowns()
        setupClickListeners()
        setupExportButtons()

        binding.buttonSecond.setOnClickListener {
            findNavController().navigate(R.id.action_SecondFragment_to_FirstFragment)
        }
    }

    private fun setupExportButtons() {
        binding.btnExportCsv.setOnClickListener {
            showInvoicePreviewDialog { format ->
                when (format) {
                    "csv" -> exportToCSV()
                }
            }
        }

        binding.btnExportExcel.setOnClickListener {
            showInvoicePreviewDialog { format ->
                when (format) {
                    "poi" -> exportToExcelWithPOI()
                    "html" -> exportToExcelHTML()
                }
            }
        }
    }

    // ==================== INVOICE PREVIEW DIALOG ====================
    private fun showInvoicePreviewDialog(onExport: (String) -> Unit) {
        // Validate selections first
        if (selectedClient == null) {
            Toast.makeText(requireContext(), "Please select a client", Toast.LENGTH_SHORT).show()
            return
        }

        if (selectedProduct == null) {
            Toast.makeText(requireContext(), "Please select a product", Toast.LENGTH_SHORT).show()
            return
        }

        val dialogView = LayoutInflater.from(requireContext()).inflate(R.layout.dialog_invoice_preview, null)

        // Find all EditText fields
        val etClientName = dialogView.findViewById<TextInputEditText>(R.id.etClientName)
        val etInvoiceDate = dialogView.findViewById<TextInputEditText>(R.id.etInvoiceDate)
        val etCGST = dialogView.findViewById<TextInputEditText>(R.id.etCGST)
        val etSGST = dialogView.findViewById<TextInputEditText>(R.id.etSGST)
        val etIGST = dialogView.findViewById<TextInputEditText>(R.id.etIGST)
        val switchReverseCharge = dialogView.findViewById<Switch>(R.id.switchReverseCharge)

        // Product details (read-only)
        val tvProductName = dialogView.findViewById<TextView>(R.id.tvProductName)
        val tvHSNCode = dialogView.findViewById<TextView>(R.id.tvHSNCode)
        val tvQuantity = dialogView.findViewById<TextView>(R.id.tvQuantity)
        val tvRate = dialogView.findViewById<TextView>(R.id.tvRate)
        val tvAmount = dialogView.findViewById<TextView>(R.id.tvAmount)

        // Totals (calculated)
        val tvSubtotal = dialogView.findViewById<TextView>(R.id.tvSubtotal)
        val tvCGSTAmount = dialogView.findViewById<TextView>(R.id.tvCGSTAmount)
        val tvSGSTAmount = dialogView.findViewById<TextView>(R.id.tvSGSTAmount)
        val tvIGSTAmount = dialogView.findViewById<TextView>(R.id.tvIGSTAmount)
        val tvGrandTotal = dialogView.findViewById<TextView>(R.id.tvGrandTotal)

        // Set default values
        etClientName.setText(selectedClient!!.name)
        etInvoiceDate.setText(invoiceDate)
        etCGST.setText(cgstRate.toString())
        etSGST.setText(sgstRate.toString())
        etIGST.setText(igstRate.toString())
        switchReverseCharge.isChecked = isReverseCharge

        // Set product details
        tvProductName.text = selectedProduct!!.name
        tvHSNCode.text = selectedProduct!!.hsnCode
        tvQuantity.text = selectedProduct!!.quantity.toString()
        tvRate.text = "₹ ${selectedProduct!!.price}"

        val subtotal = selectedProduct!!.quantity * selectedProduct!!.price
        tvAmount.text = "₹ ${String.format("%.2f", subtotal)}"

        // Calculate and display totals
        val calculateTotals = {
            try {
                val cgst = etCGST.text.toString().toDoubleOrNull() ?: 0.0
                val sgst = etSGST.text.toString().toDoubleOrNull() ?: 0.0
                val igst = etIGST.text.toString().toDoubleOrNull() ?: 0.0

                val subtotalValue = selectedProduct!!.quantity * selectedProduct!!.price
                val cgstAmount = subtotalValue * (cgst / 100)
                val sgstAmount = subtotalValue * (sgst / 100)
                val igstAmount = subtotalValue * (igst / 100)
                val grandTotal = subtotalValue + cgstAmount + sgstAmount + igstAmount

                tvSubtotal.text = "₹ ${String.format("%.2f", subtotalValue)}"
                tvCGSTAmount.text = "₹ ${String.format("%.2f", cgstAmount)}"
                tvSGSTAmount.text = "₹ ${String.format("%.2f", sgstAmount)}"
                tvIGSTAmount.text = "₹ ${String.format("%.2f", igstAmount)}"
                tvGrandTotal.text = "₹ ${String.format("%.2f", grandTotal)}"
            } catch (e: Exception) {
                Log.e(TAG, "Error calculating totals", e)
            }
        }

        // Initial calculation
        calculateTotals()

        // Update totals when tax rates change
        etCGST.addTextChangedListener(object : android.text.TextWatcher {
            override fun afterTextChanged(s: android.text.Editable?) { calculateTotals() }
            override fun beforeTextChanged(s: CharSequence?, start: Int, count: Int, after: Int) {}
            override fun onTextChanged(s: CharSequence?, start: Int, before: Int, count: Int) {}
        })

        etSGST.addTextChangedListener(object : android.text.TextWatcher {
            override fun afterTextChanged(s: android.text.Editable?) { calculateTotals() }
            override fun beforeTextChanged(s: CharSequence?, start: Int, count: Int, after: Int) {}
            override fun onTextChanged(s: CharSequence?, start: Int, before: Int, count: Int) {}
        })

        etIGST.addTextChangedListener(object : android.text.TextWatcher {
            override fun afterTextChanged(s: android.text.Editable?) { calculateTotals() }
            override fun beforeTextChanged(s: CharSequence?, start: Int, count: Int, after: Int) {}
            override fun onTextChanged(s: CharSequence?, start: Int, before: Int, count: Int) {}
        })

        val dialog = AlertDialog.Builder(requireContext())
            .setTitle("Invoice Preview & Edit")
            .setView(dialogView)
            .setPositiveButton("Export CSV") { _, _ ->
                // Save the entered data
                selectedClient = selectedClient!!.copy(name = etClientName.text.toString())
                invoiceDate = etInvoiceDate.text.toString()
                cgstRate = etCGST.text.toString().toDoubleOrNull() ?: 9.0
                sgstRate = etSGST.text.toString().toDoubleOrNull() ?: 9.0
                igstRate = etIGST.text.toString().toDoubleOrNull() ?: 0.0
                isReverseCharge = switchReverseCharge.isChecked

                onExport("csv")
            }
            .setNeutralButton("Export Excel (POI)") { _, _ ->
                // Save the entered data
                selectedClient = selectedClient!!.copy(name = etClientName.text.toString())
                invoiceDate = etInvoiceDate.text.toString()
                cgstRate = etCGST.text.toString().toDoubleOrNull() ?: 9.0
                sgstRate = etSGST.text.toString().toDoubleOrNull() ?: 9.0
                igstRate = etIGST.text.toString().toDoubleOrNull() ?: 0.0
                isReverseCharge = switchReverseCharge.isChecked

                onExport("poi")
            }
            .setNegativeButton("Cancel", null)
            .create()

        dialog.show()

        // Optionally add HTML export button
        dialog.getButton(AlertDialog.BUTTON_NEUTRAL).setOnLongClickListener {
            selectedClient = selectedClient!!.copy(name = etClientName.text.toString())
            invoiceDate = etInvoiceDate.text.toString()
            cgstRate = etCGST.text.toString().toDoubleOrNull() ?: 9.0
            sgstRate = etSGST.text.toString().toDoubleOrNull() ?: 9.0
            igstRate = etIGST.text.toString().toDoubleOrNull() ?: 0.0
            isReverseCharge = switchReverseCharge.isChecked

            dialog.dismiss()
            onExport("html")
            true
        }
    }

    // ==================== UPDATED CSV EXPORT WITH FILLED DATA ====================
    private fun exportToCSV() {
        if (selectedClient == null || selectedProduct == null) {
            Toast.makeText(requireContext(), "Missing data", Toast.LENGTH_SHORT).show()
            return
        }

        Log.d(TAG, "Starting CSV export process...")

        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
            saveCSVUsingMediaStore()
        } else {
            if (checkStoragePermission()) {
                saveCSVLegacy()
            } else {
                requestStoragePermission()
            }
        }
    }

    private fun saveCSVUsingMediaStore() {
        try {
            val productItems = listOf(selectedProduct!!)
            val dateFormat = SimpleDateFormat("yyyyMMdd_HHmmss", Locale.getDefault())
            val timestamp = dateFormat.format(Date())
            val fileName = "Invoice_${selectedClient!!.name}_$timestamp.csv"

            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                saveCSVToDownloads(fileName, productItems)
            } else {
                saveCSVToAppDirectory(fileName, productItems)
            }
        } catch (e: Exception) {
            Log.e(TAG, "Error saving CSV", e)
            Toast.makeText(requireContext(), "Error: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }

    @androidx.annotation.RequiresApi(Build.VERSION_CODES.Q)
    private fun saveCSVToDownloads(fileName: String, productItems: List<ProductItem>) {
        try {
            val contentValues = ContentValues().apply {
                put(MediaStore.Downloads.DISPLAY_NAME, fileName)
                put(MediaStore.Downloads.MIME_TYPE, "text/csv")
                put(MediaStore.Downloads.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS)
            }

            val resolver = requireContext().contentResolver
            val uri = resolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, contentValues)

            if (uri != null) {
                val outputStream = resolver.openOutputStream(uri)
                if (outputStream != null) {
                    val success = createCSVFile(outputStream, productItems)
                    if (success) {
                        Toast.makeText(requireContext(), "CSV Invoice saved to Downloads!", Toast.LENGTH_LONG).show()
                        Log.d(TAG, "CSV file saved to Downloads: $fileName")
                    } else {
                        Toast.makeText(requireContext(), "Failed to create CSV file", Toast.LENGTH_SHORT).show()
                    }
                } else {
                    saveCSVToAppDirectory(fileName, productItems)
                }
            } else {
                saveCSVToAppDirectory(fileName, productItems)
            }
        } catch (e: Exception) {
            Log.e(TAG, "Error saving CSV to Downloads", e)
            saveCSVToAppDirectory(fileName, productItems)
        }
    }

    private fun saveCSVToAppDirectory(fileName: String, productItems: List<ProductItem>) {
        try {
            val documentsDir = File(requireContext().getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS), "Invoices")
            if (!documentsDir.exists()) {
                documentsDir.mkdirs()
            }

            val file = File(documentsDir, fileName)
            val outputStream = FileOutputStream(file)

            val success = createCSVFile(outputStream, productItems)
            if (success) {
                Toast.makeText(requireContext(), "CSV Invoice exported successfully!\nCheck app Documents folder", Toast.LENGTH_LONG).show()
                Log.d(TAG, "CSV saved to: ${file.absolutePath}")
            } else {
                Toast.makeText(requireContext(), "Failed to create CSV file", Toast.LENGTH_SHORT).show()
            }
        } catch (e: Exception) {
            Log.e(TAG, "Error saving CSV to app directory", e)
            Toast.makeText(requireContext(), "Failed to save CSV: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }

    private fun saveCSVLegacy() {
        try {
            val productItems = listOf(selectedProduct!!)
            val dateFormat = SimpleDateFormat("yyyyMMdd_HHmmss", Locale.getDefault())
            val timestamp = dateFormat.format(Date())
            val fileName = "Invoice_${selectedClient!!.name}_$timestamp.csv"

            val downloadsDir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS)
            val invoicesDir = File(downloadsDir, "Invoices")
            if (!invoicesDir.exists()) {
                invoicesDir.mkdirs()
            }

            val file = File(invoicesDir, fileName)
            val outputStream = FileOutputStream(file)

            val success = createCSVFile(outputStream, productItems)
            if (success) {
                Toast.makeText(requireContext(), "CSV Invoice saved to Downloads/Invoices!", Toast.LENGTH_LONG).show()
                Log.d(TAG, "CSV file saved: ${file.absolutePath}")
            } else {
                Toast.makeText(requireContext(), "Failed to create CSV file", Toast.LENGTH_SHORT).show()
            }
        } catch (e: Exception) {
            Log.e(TAG, "Error in legacy CSV save", e)
            Toast.makeText(requireContext(), "Error: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }

    private fun createCSVFile(outputStream: OutputStream, productItems: List<ProductItem>): Boolean {
        return try {
            Log.d(TAG, "Creating CSV file with filled data...")

            val csvContent = StringBuilder()

            // Add company header
            csvContent.append("GSTIN : 09ACLPA7672F12G\n")
            csvContent.append("Sales Office: G-17, Gr. Floor, Shiva Tower, G.T. Road, Ghaziabad.\n")
            csvContent.append("H.O. : 1-B.M. Complex, 1st Floor, G.T. Road, Ghaziabad.\n")
            csvContent.append("M.: 9911177105, 7982739105 - E-mail : pentium.mts@gmail.com\n\n")

            // Add order section with actual date
            csvContent.append("Order\n")
            csvContent.append("Date: $invoiceDate\n\n")

            // Add client name (filled from preview)
            csvContent.append("M/s: ${selectedClient!!.name}\n\n")

            // Add table headers
            csvContent.append("S. No.,Description of Good,HSN code,Qty,Rate,Amount\n")

            // Add product data
            productItems.forEachIndexed { index, product ->
                val amount = product.quantity * product.price
                csvContent.append("${index + 1},${escapeCsvField(product.name)},${escapeCsvField(product.hsnCode)},${product.quantity},${product.price},$amount\n")
            }

            // Add empty lines
            csvContent.append("\n\n\n\n\n\n\n")

            // Calculate totals with user-entered rates
            val subtotal = productItems.sumOf { it.quantity * it.price }
            csvContent.append("Total Amount Before Tax,$subtotal\n")

            val cgstAmount = subtotal * (cgstRate / 100)
            val sgstAmount = subtotal * (sgstRate / 100)
            val igstAmount = subtotal * (igstRate / 100)
            val grandTotal = subtotal + cgstAmount + sgstAmount + igstAmount

            // Add GST with actual rates
            csvContent.append("CGST @$cgstRate%,$cgstAmount\n")
            csvContent.append("SGST @$sgstRate%,$sgstAmount\n")

            if (igstRate > 0) {
                csvContent.append("IGST @$igstRate%,$igstAmount\n")
            }

            if (isReverseCharge) {
                csvContent.append("GST on Reverse Charge,Yes\n")
            }

            csvContent.append("GRAND TOTAL,$grandTotal\n\n")

            // Add footer notes
            csvContent.append("Bank Name:\n")
            csvContent.append("Account No:\n")
            csvContent.append("IFSC:\n\n")
            csvContent.append("1. E. & O.E.\n")
            csvContent.append("2. No physical Electrical Damage is covered under warranty.\n")

            // Write to output stream
            outputStream.write(csvContent.toString().toByteArray())
            outputStream.close()

            Log.d(TAG, "CSV file created successfully with all data filled")
            true
        } catch (e: Exception) {
            Log.e(TAG, "Error creating CSV file", e)
            false
        }
    }
    private fun escapeCsvField(field: String): String {
        return if (field.contains(",") || field.contains("\"") || field.contains("\n")) {
            "\"" + field.replace("\"", "\"\"") + "\""
        } else {
            field
        }
    }

    // ==================== PLACEHOLDER METHODS FOR EXCEL EXPORTS ====================
    // (Add your Apache POI and HTML export methods here from previous artifacts)

    //    private fun exportToExcelWithPOI() {
    //        Toast.makeText(requireContext(), "Apache POI Export - Add implementation from artifact 1", Toast.LENGTH_SHORT).show()
    //
    //    }

// ----------------------------------------
    private fun exportToExcelWithPOI() {
        // Validate selections
        if (selectedClient == null) {
            Toast.makeText(requireContext(), "Please select a client", Toast.LENGTH_SHORT).show()
            return
        }

        if (selectedProduct == null) {
            Toast.makeText(requireContext(), "Please select a product", Toast.LENGTH_SHORT).show()
            return
        }

        try {
            val workbook = XSSFWorkbook()
            val sheet = workbook.createSheet("Tax Invoice")

            // Create styles
            val headerStyle = createHeaderStyle(workbook)
            val titleStyle = createTitleStyle(workbook)
            val boldStyle = createBoldStyle(workbook)
            val normalStyle = createNormalStyle(workbook)
            val borderStyle = createBorderStyle(workbook)
            val rightAlignStyle = createRightAlignStyle(workbook)
            val centerStyle = createCenterStyle(workbook)
            val dottedStyle = createDottedStyle(workbook)

            // Set column widths (in units of 1/256th of a character width)
            sheet.setColumnWidth(0, 1500)  // A - narrow
            sheet.setColumnWidth(1, 2000)  // B - S.No
            sheet.setColumnWidth(2, 8000)  // C - Description
            sheet.setColumnWidth(3, 4000)  // D - HSN code
            sheet.setColumnWidth(4, 2500)  // E - Qty
            sheet.setColumnWidth(5, 3000)  // F - Rate
            sheet.setColumnWidth(6, 3500)  // G - Amount

            var currentRow = 0

            // Row 0: GSTIN and TAX INVOICE
            val row0 = sheet.createRow(currentRow++)
            row0.height = 400

            val gstinCell = row0.createCell(1)
            gstinCell.setCellValue("GSTIN : 09ACLPA7672F1ZG")
            gstinCell.cellStyle = boldStyle as XSSFCellStyle?
            sheet.addMergedRegion(CellRangeAddress(0, 0, 1, 2))

            val taxInvoiceCell = row0.createCell(3)
            taxInvoiceCell.setCellValue("TAX INVOICE")
            taxInvoiceCell.cellStyle = titleStyle as XSSFCellStyle?
            sheet.addMergedRegion(CellRangeAddress(0, 0, 3, 4))

            // Row 1-4: Logo area (merged cells)
            for (i in 1..4) {
                sheet.createRow(currentRow++)
            }
            sheet.addMergedRegion(CellRangeAddress(1, 4, 1, 6))
            val logoRow = sheet.getRow(2)
            val logoCell = logoRow.createCell(1)
            logoCell.setCellValue("PENTIUM (India)")
            logoCell.cellStyle = createLogoStyle(workbook) as XSSFCellStyle?

            // Row 5: Sales Office
            val row5 = sheet.createRow(currentRow++)
            val salesOfficeCell = row5.createCell(1)
            salesOfficeCell.setCellValue("Sales Office: G-17, Gr. Floor, Shiva Tower, G.T. Road, Ghaziabad.")
            salesOfficeCell.cellStyle = centerStyle as XSSFCellStyle?
            sheet.addMergedRegion(CellRangeAddress(5, 5, 1, 6))

            // Row 6: H.O.
            val row6 = sheet.createRow(currentRow++)
            val hoCell = row6.createCell(1)
            hoCell.setCellValue("H.O. : 1-B.M. Complex, 1st Floor, G.T. Road, Ghaziabad.")
            hoCell.cellStyle = centerStyle
            sheet.addMergedRegion(CellRangeAddress(6, 6, 1, 6))

            // Row 7: Contact info
            val row7 = sheet.createRow(currentRow++)
            val contactCell = row7.createCell(1)
            contactCell.setCellValue("M.: 9911177105, 7982739105 - E-mail : pentium.mts@gmail.com")
            contactCell.cellStyle = centerStyle
            sheet.addMergedRegion(CellRangeAddress(7, 7, 1, 6))

            // Add border to header section
            addBorderToRegion(sheet, 0, 7, 1, 6, workbook)

            // Row 8: Order and Date
    //        val row8 = sheet.createRow(currentRow++)
    //        row8.height = 400
    //
    //        val orderCell = row8.createCell(1)
    //        orderCell.setCellValue("Order")
    //        orderCell.cellStyle = boldStyle as XSSFCellStyle?
    //
    //        val dateCell = row8.createCell(3)
    //        // val dateFormat = SimpleDateFormat("dd/MM/yyyy", Locale.getDefault())
    //        dateCell.setCellValue("Date.......................")
    //        dateCell.cellStyle = normalStyle as XSSFCellStyle?
    //        sheet.addMergedRegion(CellRangeAddress(8, 8, 3, 6))
    //
    //        addBorderToRegion(sheet, 8, 8, 1, 6, workbook)
            // Row 8: Order and Date
            val row8 = sheet.createRow(currentRow++)
            row8.height = 400

            val orderCell = row8.createCell(1)
            orderCell.setCellValue("Order")
            orderCell.cellStyle = boldStyle

            val dateCell = row8.createCell(3)
    // Use the actual invoice date from the preview dialog
            dateCell.setCellValue("Date: ....$invoiceDate...................................................................") // This will show "Date: 25/11/2023" or whatever date was entered
            dateCell.cellStyle = normalStyle as XSSFCellStyle?
            sheet.addMergedRegion(CellRangeAddress(8, 8, 3, 6))

            // Row 9: M/s (Client name)
    //        val row9 = sheet.createRow(currentRow++)
    //        row9.height = 400
    //
    //        val msCell = row9.createCell(1)
    //        msCell.setCellValue("M/s.................................................................................................")
    //        msCell.cellStyle = dottedStyle as XSSFCellStyle?
    //        sheet.addMergedRegion(CellRangeAddress(9, 9, 1, 6))
    //
    //        addBorderToRegion(sheet, 9, 9, 1, 6, workbook)
            // Row 9: M/s (Client name) - FILLED WITH ACTUAL DATA
            val row9 = sheet.createRow(currentRow++)
            row9.height = 400

            val msCell = row9.createCell(1)
            msCell.setCellValue("M/s: ..${selectedClient!!.name}.............................................................................") // ACTUAL CLIENT NAME
            msCell.cellStyle = normalStyle // Changed from dottedStyle to normalStyle
            sheet.addMergedRegion(CellRangeAddress(9, 9, 1, 6))

            // Row 10: Dotted line
            val row10 = sheet.createRow(currentRow++)
            val dottedCell = row10.createCell(1)
            dottedCell.setCellValue(".................................................................................................")
            dottedCell.cellStyle = dottedStyle as XSSFCellStyle?
            sheet.addMergedRegion(CellRangeAddress(10, 10, 1, 6))

            addBorderToRegion(sheet, 10, 10, 1, 6, workbook)

            // Row 11: Table Header
            val row11 = sheet.createRow(currentRow++)
            row11.height = 500

            val headers = listOf("S. No.", "Desciption of Good", "HSN code", "Qty", "Rate", "Amount")
            val headerCells = listOf(1, 2, 3, 4, 5, 6)

            headers.forEachIndexed { index, header ->
                val cell = row11.createCell(headerCells[index])
                cell.setCellValue(header)
                cell.cellStyle = createTableHeaderStyle(workbook) as XSSFCellStyle?
            }

            // Rows 12-23: Product rows (empty with borders)
            val productStartRow = currentRow
            for (i in 0..11) {
                val row = sheet.createRow(currentRow++)
                row.height = 400

                for (col in 1..6) {
                    val cell = row.createCell(col)
                    cell.cellStyle = borderStyle as XSSFCellStyle?
                }
            }

            // Add actual product data if available
            if (selectedProduct != null) {
                val productRow = sheet.getRow(productStartRow)

                productRow.getCell(1).setCellValue("1")
                productRow.getCell(2).setCellValue(selectedProduct!!.name)
                productRow.getCell(3).setCellValue(selectedProduct!!.hsnCode)
                productRow.getCell(4).setCellValue(selectedProduct!!.quantity.toDouble())
                productRow.getCell(5).setCellValue(selectedProduct!!.price.toDouble())

                val amount = selectedProduct!!.quantity * selectedProduct!!.price
                productRow.getCell(6).setCellValue(amount.toDouble())
            }

            // Row 24: Amount in words
            val row24 = sheet.createRow(currentRow++)
            row24.height = 400

            val amountWordsCell = row24.createCell(1)
            amountWordsCell.setCellValue("Amount in words .................................................................")
            amountWordsCell.cellStyle = normalStyle
            sheet.addMergedRegion(CellRangeAddress(24, 24, 1, 4))

            val totalBeforeTaxCell = row24.createCell(5)
            totalBeforeTaxCell.setCellValue("Total Amount Before Tax")
            totalBeforeTaxCell.cellStyle = normalStyle
            sheet.addMergedRegion(CellRangeAddress(24, 24, 5, 6))

            addBorderToRegion(sheet, 24, 24, 1, 6, workbook)

            // Row 25: CGST
    //        val row25 = sheet.createRow(currentRow++)
    //        row25.height = 400
    //
    //        val cgstDotCell = row25.createCell(1)
    //        cgstDotCell.setCellValue(".................................................................................................")
    //        cgstDotCell.cellStyle = dottedStyle
    //        sheet.addMergedRegion(CellRangeAddress(25, 25, 1, 4))
    //
    //        val cgstCell = row25.createCell(5)
    //        cgstCell.setCellValue("CGST @")
    //        cgstCell.cellStyle = normalStyle
    //        sheet.addMergedRegion(CellRangeAddress(25, 25, 5, 6))
    //
    //        addBorderToRegion(sheet, 25, 25, 1, 6, workbook)
            //***************Row 25: CGST*************
            // Row 25: CGST - FILLED WITH ACTUAL DATA
            val row25 = sheet.createRow(currentRow++)
            row25.height = 400

    // Calculate amounts
            val subtotal = selectedProduct!!.quantity * selectedProduct!!.price
            val cgstAmount = subtotal * (cgstRate / 100)
            val sgstAmount = subtotal * (sgstRate / 100)
            val igstAmount = subtotal * (igstRate / 100)

           // val amountWordsCell = row25.createCell(1)
            amountWordsCell.setCellValue("Amount in words: ${convertNumberToWords(subtotal.toInt())}")
            amountWordsCell.cellStyle = normalStyle
            sheet.addMergedRegion(CellRangeAddress(25, 25, 1, 4))

            val cgstCell = row25.createCell(5)
            cgstCell.setCellValue("CGST @${cgstRate}%")
            cgstCell.cellStyle = normalStyle

            val cgstValueCell = row25.createCell(6)
            cgstValueCell.setCellValue(cgstAmount)
            cgstValueCell.cellStyle = borderStyle as XSSFCellStyle?
            //**************Row 25: CGST END**************

            // Row 26: SGST
    //        val row26 = sheet.createRow(currentRow++)
    //        row26.height = 400
    //
    //        val sgstCell = row26.createCell(5)
    //        sgstCell.setCellValue("SGST @")
    //        sgstCell.cellStyle = normalStyle
    //        sheet.addMergedRegion(CellRangeAddress(26, 26, 5, 6))
    //
    //        addBorderToRegion(sheet, 26, 26, 5, 6, workbook)
            //*******// Row 26: SGST - FILLED WITH ACTUAL DATA START
            // Row 26: SGST - FILLED WITH ACTUAL DATA
            val row26 = sheet.createRow(currentRow++)
            row26.height = 400

            val sgstCell = row26.createCell(5)
            sgstCell.setCellValue("SGST @${sgstRate}%")
            sgstCell.cellStyle = normalStyle

            val sgstValueCell = row26.createCell(6)
            sgstValueCell.setCellValue(sgstAmount)
            sgstValueCell.cellStyle = borderStyle

            //*******// Row 26: SGST - FILLED WITH ACTUAL DATA END

            // Row 27: IGST
    //        val row27 = sheet.createRow(currentRow++)
    //        row27.height = 400
    //
    //        val igstCell = row27.createCell(5)
    //        igstCell.setCellValue("IGST @")
    //        igstCell.cellStyle = normalStyle
    //        sheet.addMergedRegion(CellRangeAddress(27, 27, 5, 6))
    //
    //        addBorderToRegion(sheet, 27, 27, 5, 6, workbook)
            //************// Row 27: IGST - FILLED WITH ACTUAL DATA (only if rate > 0) START
            // Row 27: IGST - FILLED WITH ACTUAL DATA (only if rate > 0)
            if (igstRate > 0) {
                val row27 = sheet.createRow(currentRow++)
                row27.height = 400

                val igstCell = row27.createCell(5)
                igstCell.setCellValue("IGST @${igstRate}%")
                igstCell.cellStyle = normalStyle

                val igstValueCell = row27.createCell(6)
                igstValueCell.setCellValue(igstAmount)
                igstValueCell.cellStyle = borderStyle
            }
            //************// Row 27: IGST - FILLED WITH ACTUAL DATA (only if rate > 0) END

            // Row 28: GST on Reverse Charge
            val row28 = sheet.createRow(currentRow++)
            row28.height = 400

            val reverseChargeCell = row28.createCell(5)
            reverseChargeCell.setCellValue("GST on Reverse Charge")
            reverseChargeCell.cellStyle = normalStyle
            sheet.addMergedRegion(CellRangeAddress(28, 28, 5, 6))

            addBorderToRegion(sheet, 28, 28, 5, 6, workbook)

            // Row 29: Bank Name and Grand Total
    //        val row29 = sheet.createRow(currentRow++)
    //        row29.height = 400
    //
    //        val bankNameCell = row29.createCell(1)
    //        bankNameCell.setCellValue("Bank Name:")
    //        bankNameCell.cellStyle = normalStyle
    //        sheet.addMergedRegion(CellRangeAddress(29, 29, 1, 4))
    //
    //        val grandTotalCell = row29.createCell(5)
    //        grandTotalCell.setCellValue("GRAND TOTAL")
    //        grandTotalCell.cellStyle = createGrandTotalStyle(workbook) as XSSFCellStyle?
    //        sheet.addMergedRegion(CellRangeAddress(29, 29, 5, 6))
            //********// Row 29: Bank Name and Grand Total - FILLED WITH ACTUAL DATA START
            // Row 29: Bank Name and Grand Total - FILLED WITH ACTUAL DATA
            val row29 = sheet.createRow(currentRow++)
            row29.height = 400

            val bankNameCell = row29.createCell(1)
            bankNameCell.setCellValue("Bank Name:")
            bankNameCell.cellStyle = normalStyle
            sheet.addMergedRegion(CellRangeAddress(29, 29, 1, 4))

            val grandTotalCell = row29.createCell(5)
            val grandTotal = subtotal + cgstAmount + sgstAmount + igstAmount
            grandTotalCell.setCellValue("GRAND TOTAL: ₹${String.format("%.2f", grandTotal)}")
            grandTotalCell.cellStyle = createGrandTotalStyle(workbook) as XSSFCellStyle?
            sheet.addMergedRegion(CellRangeAddress(29, 29, 5, 6))
            //********// Row 29: Bank Name and Grand Total - FILLED WITH ACTUAL DATA END

            addBorderToRegion(sheet, 29, 29, 1, 6, workbook)

            // Row 30: Account No
            val row30 = sheet.createRow(currentRow++)
            row30.height = 400

            val accountCell = row30.createCell(1)
            accountCell.setCellValue("Account No:")
            accountCell.cellStyle = normalStyle
            sheet.addMergedRegion(CellRangeAddress(30, 30, 1, 4))

            addBorderToRegion(sheet, 30, 30, 1, 4, workbook)

            // Row 31: IFSC
            val row31 = sheet.createRow(currentRow++)
            row31.height = 400

            val ifscCell = row31.createCell(1)
            ifscCell.setCellValue("IFSC           :")
            ifscCell.cellStyle = normalStyle
            sheet.addMergedRegion(CellRangeAddress(31, 31, 1, 4))

            addBorderToRegion(sheet, 31, 31, 1, 4, workbook)

            // Row 32: E. & O.E.
            val row32 = sheet.createRow(currentRow++)
            row32.height = 400

            val eoeCell = row32.createCell(1)
            eoeCell.setCellValue("1. E. & O.E.")
            eoeCell.cellStyle = normalStyle
            sheet.addMergedRegion(CellRangeAddress(32, 32, 1, 6))

            // Row 33: Warranty note
            val row33 = sheet.createRow(currentRow++)
            row33.height = 400

            val warrantyCell = row33.createCell(1)
            warrantyCell.setCellValue("2. No physical Electrical Damage is covered under warranty.")
            warrantyCell.cellStyle = normalStyle
            sheet.addMergedRegion(CellRangeAddress(33, 33, 1, 6))

            // Save file
            val dateFormat = SimpleDateFormat("yyyyMMdd_HHmmss", Locale.getDefault())
            val timestamp = dateFormat.format(Date())
            val fileName = "Invoice_${selectedClient!!.name}_$timestamp.xlsx"

            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                saveWorkbookToDownloads(workbook, fileName)
            } else {
                saveWorkbookToAppDirectory(workbook, fileName)
            }

        } catch (e: Exception) {
            Log.e(TAG, "Error creating Excel with POI", e)
            Toast.makeText(requireContext(), "Error: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }

    private fun convertNumberToWords(number: Int): String {
        val units = arrayOf("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine")
        val teens = arrayOf("Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
        val tens = arrayOf("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")

        return when {
            number == 0 -> "Zero"
            number < 10 -> units[number]
            number < 20 -> teens[number - 10]
            number < 100 -> "${tens[number / 10]} ${units[number % 10]}".trim()
            number < 1000 -> "${units[number / 100]} Hundred ${convertNumberToWords(number % 100)}".trim()
            else -> "Rupees"
        } + " Only"
    }

    // Style creation methods
    private fun createHeaderStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        val font = workbook.createFont()
        font.bold = true
        font.fontHeightInPoints = 11
        style.setFont(font)
        style.alignment = HorizontalAlignment.LEFT
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }

    private fun createTitleStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        val font = workbook.createFont()
        font.bold = true
        font.fontHeightInPoints = 14
        font.underline = Font.U_SINGLE
        style.setFont(font)
        style.alignment = HorizontalAlignment.CENTER
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }

    private fun createBoldStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        val font = workbook.createFont()
        font.bold = true
        font.fontHeightInPoints = 10
        style.setFont(font)
        style.alignment = HorizontalAlignment.LEFT
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }

    private fun createNormalStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        val font = workbook.createFont()
        font.fontHeightInPoints = 10
        style.setFont(font)
        style.alignment = HorizontalAlignment.LEFT
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }

    private fun createBorderStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        style.borderTop = BorderStyle.THIN
        style.borderBottom = BorderStyle.THIN
        style.borderLeft = BorderStyle.THIN
        style.borderRight = BorderStyle.THIN
        style.alignment = HorizontalAlignment.CENTER
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }

    private fun createTableHeaderStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        val font = workbook.createFont()
        font.bold = true
        font.fontHeightInPoints = 10
        style.setFont(font)
        style.borderTop = BorderStyle.THIN
        style.borderBottom = BorderStyle.THIN
        style.borderLeft = BorderStyle.THIN
        style.borderRight = BorderStyle.THIN
        style.alignment = HorizontalAlignment.CENTER
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }

    private fun createRightAlignStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        style.alignment = HorizontalAlignment.RIGHT
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }

    private fun createCenterStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        val font = workbook.createFont()
        font.fontHeightInPoints = 10
        style.setFont(font)
        style.alignment = HorizontalAlignment.CENTER
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }

    private fun createDottedStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        style.alignment = HorizontalAlignment.LEFT
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }

    private fun createLogoStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        val font = workbook.createFont()
        font.bold = true
        font.fontHeightInPoints = 20
        style.setFont(font)
        style.alignment = HorizontalAlignment.CENTER
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }

    private fun createGrandTotalStyle(workbook: Workbook): CellStyle {
        val style = workbook.createCellStyle()
        val font = workbook.createFont()
        font.bold = true
        font.fontHeightInPoints = 11
        style.setFont(font)
        style.borderTop = BorderStyle.THIN
        style.borderBottom = BorderStyle.THIN
        style.borderLeft = BorderStyle.THIN
        style.borderRight = BorderStyle.THIN
        style.alignment = HorizontalAlignment.CENTER
        style.verticalAlignment = VerticalAlignment.CENTER
        return style
    }

    // Helper method to add borders to a region
    private fun addBorderToRegion(sheet: Sheet, startRow: Int, endRow: Int, startCol: Int, endCol: Int, workbook: Workbook) {
        val borderStyle = BorderStyle.THIN

        for (row in startRow..endRow) {
            val currentRow = sheet.getRow(row) ?: sheet.createRow(row)

            for (col in startCol..endCol) {
                val cell = currentRow.getCell(col) ?: currentRow.createCell(col)
                val style = cell.cellStyle ?: workbook.createCellStyle()

                // Top border
                if (row == startRow) {
                    style.borderTop = borderStyle
                }
                // Bottom border
                if (row == endRow) {
                    style.borderBottom = borderStyle
                }
                // Left border
                if (col == startCol) {
                    style.borderLeft = borderStyle
                }
                // Right border
                if (col == endCol) {
                    style.borderRight = borderStyle
                }

                cell.cellStyle = style
            }
        }
    }

    // Save workbook methods
    @androidx.annotation.RequiresApi(Build.VERSION_CODES.Q)
    private fun saveWorkbookToDownloads(workbook: Workbook, fileName: String) {
        try {
            val contentValues = ContentValues().apply {
                put(MediaStore.Downloads.DISPLAY_NAME, fileName)
                put(MediaStore.Downloads.MIME_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                put(MediaStore.Downloads.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS)
            }

            val resolver = requireContext().contentResolver
            val uri = resolver.insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, contentValues)

            if (uri != null) {
                resolver.openOutputStream(uri)?.use { outputStream ->
                    workbook.write(outputStream)
                    Toast.makeText(requireContext(), "Excel Invoice saved to Downloads!", Toast.LENGTH_LONG).show()
                    Log.d(TAG, "Excel file saved to Downloads: $fileName")
                }
            }

            workbook.close()
        } catch (e: Exception) {
            Log.e(TAG, "Error saving Excel to Downloads", e)
            saveWorkbookToAppDirectory(workbook, fileName)
        }
    }

    private fun saveWorkbookToAppDirectory(workbook: Workbook, fileName: String) {
        try {
            val documentsDir = File(requireContext().getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS), "Invoices")
            if (!documentsDir.exists()) {
                documentsDir.mkdirs()
            }

            val file = File(documentsDir, fileName)
            FileOutputStream(file).use { outputStream ->
                workbook.write(outputStream)
                Toast.makeText(requireContext(), "Excel Invoice saved!\nLocation: ${file.absolutePath}", Toast.LENGTH_LONG).show()
                Log.d(TAG, "Excel saved to: ${file.absolutePath}")
            }

            workbook.close()
        } catch (e: Exception) {
            Log.e(TAG, "Error saving Excel to app directory", e)
            Toast.makeText(requireContext(), "Failed to save Excel: ${e.message}", Toast.LENGTH_LONG).show()
        }
    }
// ----------------------------------------


// ----------------------------------------

    private fun exportToExcelHTML() {
        Toast.makeText(requireContext(), "HTML Export - Add implementation from artifact 2", Toast.LENGTH_SHORT).show()
    }

    // ==================== PERMISSION HANDLING ====================
    private fun checkStoragePermission(): Boolean {
        return ContextCompat.checkSelfPermission(
            requireContext(),
            Manifest.permission.WRITE_EXTERNAL_STORAGE
        ) == PackageManager.PERMISSION_GRANTED
    }

    private fun requestStoragePermission() {
        requestPermissions(
            arrayOf(Manifest.permission.WRITE_EXTERNAL_STORAGE),
            STORAGE_PERMISSION_CODE
        )
    }

    override fun onRequestPermissionsResult(
        requestCode: Int,
        permissions: Array<out String>,
        grantResults: IntArray
    ) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        if (requestCode == STORAGE_PERMISSION_CODE) {
            if (grantResults.isNotEmpty() && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
                exportToCSV()
            } else {
                Toast.makeText(requireContext(), "Storage permission denied", Toast.LENGTH_SHORT).show()
            }
        }
    }

    // ==================== EXISTING DROPDOWN METHODS ====================
    private fun setupDropdowns() {
        val inflater = LayoutInflater.from(requireContext())

        clientDropdownView = inflater.inflate(R.layout.dropdown_layout, binding.dropdownContainer, false)
        productDropdownView = inflater.inflate(R.layout.dropdown_layout, binding.dropdownContainer, false)

        val layoutParams = FrameLayout.LayoutParams(
            FrameLayout.LayoutParams.MATCH_PARENT,
            FrameLayout.LayoutParams.WRAP_CONTENT
        )
        clientDropdownView.layoutParams = layoutParams
        productDropdownView.layoutParams = layoutParams

        clientDropdownView.visibility = View.GONE
        productDropdownView.visibility = View.GONE

        binding.dropdownContainer.addView(clientDropdownView)
        binding.dropdownContainer.addView(productDropdownView)
    }

    private fun setupClickListeners() {
        binding.clientNameLayout.setOnClickListener {
            toggleClientDropdown()
        }

        binding.descriptionOfGoodsLayout.setOnClickListener {
            toggleProductDropdown()
        }

        binding.clientArrow.setOnClickListener {
            toggleClientDropdown()
        }

        binding.productArrow.setOnClickListener {
            toggleProductDropdown()
        }
    }

    private fun toggleClientDropdown() {
        if (isClientDropdownOpen) {
            hideClientDropdown()
        } else {
            showClientDropdown()
            hideProductDropdown()
        }
    }

    private fun toggleProductDropdown() {
        if (isProductDropdownOpen) {
            hideProductDropdown()
        } else {
            showProductDropdown()
            hideClientDropdown()
        }
    }

    private fun showClientDropdown() {
        if (clientList.isEmpty()) {
            Toast.makeText(requireContext(), "No clients available", Toast.LENGTH_SHORT).show()
            return
        }

        val clientNames = clientList.map { it.name }
        val recyclerView = clientDropdownView.findViewById<RecyclerView>(R.id.recyclerViewDropdown)
        recyclerView.layoutManager = LinearLayoutManager(requireContext())
        recyclerView.adapter = DropdownAdapter(clientNames) { selectedClientName ->
            val client = clientList.find { it.name == selectedClientName }
            selectedClient = client
            binding.ClientName.text = selectedClientName
            hideClientDropdown()
        }

        setupDropdownPosition(clientDropdownView, binding.clientNameLayout)
        clientDropdownView.visibility = View.VISIBLE
        isClientDropdownOpen = true
        rotateArrow(binding.clientArrow, 180f)
    }

    private fun hideClientDropdown() {
        clientDropdownView.visibility = View.GONE
        isClientDropdownOpen = false
        rotateArrow(binding.clientArrow, 0f)
    }

    private fun showProductDropdown() {
        if (productList.isEmpty()) {
            Toast.makeText(requireContext(), "No products available", Toast.LENGTH_SHORT).show()
            return
        }

        val productNames = productList.map { it.name }
        val recyclerView = productDropdownView.findViewById<RecyclerView>(R.id.recyclerViewDropdown)
        recyclerView.layoutManager = LinearLayoutManager(requireContext())
        recyclerView.adapter = DropdownAdapter(productNames) { selectedProductName ->
            val product = productList.find { it.name == selectedProductName }
            selectedProduct = product
            binding.descriptionText.text = selectedProductName
            hideProductDropdown()
        }

        setupDropdownPosition(productDropdownView, binding.descriptionOfGoodsLayout)
        productDropdownView.visibility = View.VISIBLE
        isProductDropdownOpen = true
        rotateArrow(binding.productArrow, 180f)
    }

    private fun hideProductDropdown() {
        productDropdownView.visibility = View.GONE
        isProductDropdownOpen = false
        rotateArrow(binding.productArrow, 0f)
    }

    private fun setupDropdownPosition(dropdownView: View, anchorView: View) {
        try {
            val location = IntArray(2)
            anchorView.getLocationOnScreen(location)

            val params = dropdownView.layoutParams as? FrameLayout.LayoutParams
                ?: FrameLayout.LayoutParams(
                    FrameLayout.LayoutParams.MATCH_PARENT,
                    FrameLayout.LayoutParams.WRAP_CONTENT
                )

            params.topMargin = location[1] + anchorView.height - getStatusBarHeight()
            dropdownView.layoutParams = params
        } catch (e: Exception) {
            Log.e(TAG, "Error setting dropdown position", e)
        }
    }

    private fun getStatusBarHeight(): Int {
        var result = 0
        val resourceId = resources.getIdentifier("status_bar_height", "dimen", "android")
        if (resourceId > 0) {
            result = resources.getDimensionPixelSize(resourceId)
        }
        return result
    }

    private fun rotateArrow(imageButton: ImageButton, degrees: Float) {
        imageButton.animate().rotation(degrees).setDuration(300).start()
    }

    override fun onDestroyView() {
        super.onDestroyView()
        _binding = null
    }
}