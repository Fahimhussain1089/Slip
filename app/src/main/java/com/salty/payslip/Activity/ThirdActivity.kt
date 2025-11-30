package com.salty.payslip.Activity

import androidx.appcompat.widget.SearchView
import android.annotation.SuppressLint
import android.content.ActivityNotFoundException
import android.content.ContentValues
import android.content.Intent
import android.net.Uri
import android.os.Build
import android.os.Bundle
import android.os.Environment
import android.provider.MediaStore
import android.text.Editable
import android.text.TextWatcher
import android.util.Log
import android.view.LayoutInflater
import android.view.MenuItem
import android.view.MotionEvent
import android.view.View
import android.widget.EditText
import android.widget.FrameLayout
import android.widget.ImageButton
import android.widget.ImageView
import android.widget.LinearLayout
import android.widget.PopupMenu
import android.widget.TextView
import android.widget.Toast
import androidx.activity.enableEdgeToEdge
import androidx.activity.result.contract.ActivityResultContracts
import androidx.annotation.RequiresApi
import androidx.appcompat.app.AppCompatActivity
import androidx.appcompat.widget.AppCompatButton
import androidx.core.content.ContentProviderCompat.requireContext
import androidx.core.content.ContextCompat
import androidx.core.content.FileProvider
import androidx.lifecycle.lifecycleScope

import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import androidx.room.Room
import com.salty.payslip.Adapter.DropdownAdapter
import com.salty.payslip.Adapter.DropdownSearchableAdapter
import com.salty.payslip.Fragment.SecondFragment
import com.salty.payslip.R
import com.salty.payslip.model.Client
import com.salty.payslip.model.ProductItem
import com.salty.payslip.roomdb.AppDatabase
import com.salty.payslip.roomdb.ClientDao
import com.salty.payslip.roomdb.ProductDao
import kotlinx.coroutines.launch
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.HorizontalAlignment
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.VerticalAlignment
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileOutputStream
import java.io.InputStream
import java.text.SimpleDateFormat
import java.util.Date
import java.util.Locale

class ThirdActivity : AppCompatActivity() {

    lateinit var productContainer: LinearLayout
    lateinit var fab: ImageView
    private lateinit var btnMenu: ImageView
    private val clientList = mutableListOf<Client>()
    private val productList = mutableListOf<ProductItem>()

    private var isClientDropdownOpen = false
    private var isProductDropdownOpen = false

    private lateinit var clientDropdownView: View
    private lateinit var productDropdownView: View

    private var selectedClient: Client? = null
    private var selectedProduct: ProductItem? = null

    // For dynamic sections
    private val productSections = mutableListOf<View>()
    private var productSectionCounter = 1
    private val sectionDropdowns = mutableMapOf<View, Pair<View, View>>()
//    private val sectionSelectedClients = mutableMapOf<View, Client?>()
    private val sectionSelectedProducts = mutableMapOf<View, ProductItem?>()

    // Invoice vars
    private var invoiceDate: String = SimpleDateFormat("dd/MM/yyyy", Locale.getDefault()).format(
        Date()
    )
//    private var cgstRate: Double = 9.0
    private lateinit var cgstRate: EditText
    private lateinit var sgstRate: EditText
    private lateinit var igstRate: EditText
//    private var sgstRate: Double = 9.0
//    private var igstRate: Double = 0.0
    private var isReverseCharge: Boolean = false

    private lateinit var btnAddProduct: ImageView

    private lateinit var clientArrow: ImageView
    private lateinit var productArrow: ImageView

    private lateinit var layoutClientDropdown: LinearLayout
    private lateinit var layoutProductDropdown: LinearLayout

    private lateinit var tvClientName: TextView
    private lateinit var tvDescriptionMore: TextView

    private lateinit var dropdownContainer: FrameLayout
    private  lateinit var mainLinearLayout: LinearLayout
    private  lateinit var btn_export: AppCompatButton
    private lateinit var etQty : EditText
    private var currentlyOpenDropdown: View? = null
    private var currentlyOpenDropdownView: View? = null   // This is the KEY!
    private var isExporting = false
    // Add these bank detail variables
    private lateinit var etBankName: EditText
    private lateinit var etIfsc: EditText
    private lateinit var etAcNo: EditText

    //******** integrate the room db
    private lateinit var db: AppDatabase
    private lateinit var clientDao: ClientDao
    private lateinit var productDao: ProductDao

    private val filePickerLauncher = registerForActivityResult(ActivityResultContracts.StartActivityForResult()) { result ->
        if (result.resultCode == RESULT_OK) {
            result.data?.data?.let { uri ->
                readFileFromUri(uri)
            }
        }
    }

    companion object {
        private const val STORAGE_PERMISSION_CODE = 1001
        private const val TAG = "ThirdActivity"
    }

    @SuppressLint("MissingInflatedId")
    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        enableEdgeToEdge()
//        setContentView(R.layout.activity_third)
        setContentView(R.layout.activity_third2)

        productContainer = findViewById(R.id.product_container)
        fab = findViewById(R.id.btnAddProduct)
        //*********id find out strat ****
        btnMenu = findViewById(R.id.btn_menu)
        productContainer = findViewById(R.id.product_container)
        btnAddProduct = findViewById(R.id.btnAddProduct)
        btnMenu = findViewById(R.id.btn_menu)
        dropdownContainer = findViewById(R.id.dropdown_container)
        // Main dropdown views
        layoutClientDropdown = findViewById(R.id.layout_client_dropdown)
      //  layoutProductDropdown = findViewById(R.id.layout_product_dropdown) //by me disable the cliable property
        clientArrow = findViewById(R.id.client_arrow)
        productArrow = findViewById(R.id.product_arrow)
        tvClientName = findViewById(R.id.tv_client_name)
        tvDescriptionMore = findViewById(R.id.tv_description_more)
        mainLinearLayout = findViewById(R.id.main_linear_layout)
        btn_export = findViewById(R.id.btn_export)
         etQty = findViewById(R.id.et_Qty)
        cgstRate =  findViewById<EditText>(R.id.et_CGST)
        sgstRate =  findViewById<EditText>(R.id.et_SGST)
        igstRate =  findViewById<EditText>(R.id.et_GST)
        etBankName = findViewById(R.id.et_bankName)
        etIfsc = findViewById(R.id.et_Ifsc)
        etAcNo = findViewById(R.id.et_acNo)



        //*********id find out End  ****
        // Disable Export button at start
        btn_export.isEnabled = false
        btn_export.alpha = 0.6f  // Optional: make it look faded

        // Setup dropdowns
        setupDropdowns()
        setupClickListeners()

        // Add button
        btnAddProduct.setOnClickListener {
            addNewProductSection()
        }

        fab.setOnClickListener {
            addNewProductSection()
        }
        btn_export.setOnClickListener {
            // PREVENT MULTIPLE CLICKS — THIS IS BULLETPROOF
            if (isExporting) {
                Toast.makeText(this, "Please wait, exporting in progress...", Toast.LENGTH_SHORT).show()
                return@setOnClickListener
            }

            // Check if products exist
            if (productSections.isEmpty()) {
                Toast.makeText(this, "Please add at least one product", Toast.LENGTH_SHORT).show()
                return@setOnClickListener
            }

            // Start export
            isExporting = true
            btn_export.isEnabled = false
            btn_export.text = "Export Data"

//            exportToExcelWithPOI()
            exportToExcelWithPOI {
                // This runs when export finishes (success or fail)
                isExporting = false
                btn_export.text = "Export Data"
                btn_export.isEnabled = true
            }
        }


        // ← THIS IS THE MAGIC PART
        btnMenu.setOnClickListener { view ->
            val popup = PopupMenu(this, view)
            popup.inflate(R.menu.menu_main)

            popup.setOnMenuItemClickListener { item: MenuItem ->
                when (item.itemId) {
                    R.id.menu_upload_client -> {
                        openFilePicker("client")
                        true
                    }
                    R.id.menu_upload_product -> {
                        openFilePicker("product")
                        true
                    }
                    R.id.menu_clear_data -> {
                        clearAllData()
                        true
                    }
                    else -> false
                }
            }
            popup.show()
        }
        // Initial arrow colors (gray if empty)
        updateArrowColors()
        //------------------- For the Room db Start-----
        db = Room.databaseBuilder(
            applicationContext,
            AppDatabase::class.java,
            "app_database"
        ).allowMainThreadQueries() // For simplicity; remove later for strict async
            .build()

        clientDao = db.clientDao()
        productDao = db.productDao()

// Load initial data from DB
        lifecycleScope.launch {
            refreshLists()
        }
        //------------------- For the Room db  End -----




    }
    //********* here is the methode *************

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
    private fun setupDropdowns() {
        val inflater = layoutInflater

        clientDropdownView = inflater.inflate(R.layout.dropdown_layout, dropdownContainer, false)
        productDropdownView = inflater.inflate(R.layout.dropdown_layout, dropdownContainer, false)

        val layoutParams = FrameLayout.LayoutParams(
            FrameLayout.LayoutParams.MATCH_PARENT,
            FrameLayout.LayoutParams.WRAP_CONTENT
        )
        clientDropdownView.layoutParams = layoutParams
        productDropdownView.layoutParams = layoutParams

        clientDropdownView.visibility = View.GONE
        productDropdownView.visibility = View.GONE

        dropdownContainer.addView(clientDropdownView)
        dropdownContainer.addView(productDropdownView)
    }

    private fun setupClickListeners() {
        layoutClientDropdown.setOnClickListener {
            toggleClientDropdown()
        }
        clientArrow.setOnClickListener {
            toggleClientDropdown()
        }

//        layoutProductDropdown.setOnClickListener { //diable the cliable property by me
//            toggleProductDropdown()
//        }
//        productArrow.setOnClickListener {
//            toggleProductDropdown()
//        }//diable the cliable property by me

        // Close dropdowns on root click
        findViewById<View>(android.R.id.content).setOnClickListener {
            hideClientDropdown()
            hideProductDropdown()
            // Close section dropdowns too, similar to SecondFragment
        }
    }

//private fun showClientDropdown() {
//    if (clientList.isEmpty()) {
//        Toast.makeText(this, "No clients available", Toast.LENGTH_SHORT).show()
//        return
//    }
//
//    val clientNames = clientList.map { it.name }
//
//    setupDropdownWithSearch(clientDropdownView, clientNames) { selectedName ->
//        selectedClient = clientList.find { it.name == selectedName }
//        tvClientName.text = selectedName
//        hideClientDropdown()
//    }
//
//    showDropdownBelow(clientDropdownView, layoutClientDropdown)
//    isClientDropdownOpen = true
//    rotateArrow(clientArrow, 180f)
//}
private fun showClientDropdown() {
    if (clientList.isEmpty()) return Toast.makeText(this, "No clients available", Toast.LENGTH_SHORT).show()

    val clientNames = clientList.map { it.name }
    setupDropdownWithSearch(clientDropdownView, clientNames) { selectedName ->
        selectedClient = clientList.find { it.name == selectedName }
        tvClientName.text = selectedName
        hideClientDropdown()
    }

    showDropdownBelow(clientDropdownView, layoutClientDropdown)
    isClientDropdownOpen = true
    rotateArrow(clientArrow, 180f)
}
    private fun hideClientDropdown() {
        clientDropdownView.visibility = View.GONE
        isClientDropdownOpen = false
        rotateArrow(clientArrow, 0f)
    }
    private fun showProductDropdown() {
        if (productList.isEmpty()) {
            Toast.makeText(this, "No products available", Toast.LENGTH_SHORT).show()
            return
        }

        val productNames = productList.map { it.name }

        setupDropdownWithSearch(productDropdownView, productNames) { selectedName ->
            selectedProduct = productList.find { it.name == selectedName }
            tvDescriptionMore.text = selectedName
            hideProductDropdown()
        }

        showDropdownBelow(productDropdownView, layoutProductDropdown)//by me diable the clicable property
        showDropdownBelow(productDropdownView, findViewById(R.id.section_product_more)) // or your anchor
        isProductDropdownOpen = true
        rotateArrow(productArrow, 180f)
    }
    private fun hideProductDropdown() {
        productDropdownView.visibility = View.GONE
        isProductDropdownOpen = false
        rotateArrow(productArrow, 0f)
    }
//private fun showDropdownBelow(dropdownView: View, anchorView: View) {
//    dropdownView.post {
//        val anchorLocation = IntArray(2)
//        anchorView.getLocationInWindow(anchorLocation)
//
//        val containerLocation = IntArray(2)
//        dropdownContainer.getLocationInWindow(containerLocation)
//
//        val x = anchorLocation[0] - containerLocation[0]
//        val y = anchorLocation[1] + anchorView.height - containerLocation[1]
//
//        val params = dropdownView.layoutParams as FrameLayout.LayoutParams
//        params.leftMargin = x
//        params.topMargin = y
//        params.width = anchorView.width
//        dropdownView.layoutParams = params
//
//        dropdownView.visibility = View.VISIBLE
//        currentlyOpenDropdown = dropdownView
//    }
//}
private fun showDropdownBelow(dropdownView: View, anchorView: View) {
    // Close ANY previously open dropdown (global OR section)
    currentlyOpenDropdownView?.let { it.visibility = View.GONE }

    dropdownView.post {
        val anchorLocation = IntArray(2)
        anchorView.getLocationInWindow(anchorLocation)
        val containerLocation = IntArray(2)
        dropdownContainer.getLocationInWindow(containerLocation)

        val x = anchorLocation[0] - containerLocation[0]
        val y = anchorLocation[1] + anchorView.height - containerLocation[1]

        val params = dropdownView.layoutParams as FrameLayout.LayoutParams
        params.leftMargin = x
        params.topMargin = y
        params.width = anchorView.width
        dropdownView.layoutParams = params

        dropdownView.visibility = View.VISIBLE
        currentlyOpenDropdownView = dropdownView   // Remember which one is open
    }
}

    private fun getStatusBarHeight(): Int {
        val resourceId = resources.getIdentifier("status_bar_height", "dimen", "android")
        return if (resourceId > 0) resources.getDimensionPixelSize(resourceId) else 0
    }

    private fun rotateArrow(imageView: ImageView, degrees: Float) {
        imageView.animate().rotation(degrees).setDuration(300).start()
    }

    private fun updateArrowColors() {
        val gray = ContextCompat.getColor(this, android.R.color.darker_gray)
        val black = ContextCompat.getColor(this, android.R.color.black)

        clientArrow.setColorFilter(if (clientList.isEmpty()) gray else black)
        productArrow.setColorFilter(if (productList.isEmpty()) gray else black)
    }

    // File upload and parsing (same as SecondFragment)
    private fun openFilePicker(fileType: String) {
        val intent = Intent(Intent.ACTION_GET_CONTENT).apply {
            type = "*/*"
            addCategory(Intent.CATEGORY_OPENABLE)
            val mimeTypes = arrayOf("text/csv", "text/comma-separated-values")
            putExtra(Intent.EXTRA_MIME_TYPES, mimeTypes)
        }
        filePickerLauncher.launch(Intent.createChooser(intent, "Select $fileType CSV"))
    }

    private fun readFileFromUri(uri: Uri) {
        try {
            val fileName = getFileName(uri)
            val fileType = detectFileType(fileName)
            contentResolver.openInputStream(uri)?.use { inputStream ->
                readFileContent(inputStream, fileType, fileName)
            }
            updateArrowColors()  // Update arrows after upload
        } catch (e: Exception) {
            e.printStackTrace()
            showMessage("Error reading file: ${e.message}")
        }
    }
//    private fun clearAllData() {
//        clientList.clear()
//        productList.clear()
//        selectedClient = null
//        selectedProduct = null
//        sectionSelectedClients.clear()
//        sectionSelectedProducts.clear()
//        tvClientName.text = "Client Name (Dropdown)"
//        tvDescriptionMore.text = "Description of Goods (Dropdown)"
//        productContainer.removeAllViews()
//        updateArrowColors()
//        Toast.makeText(this, "All data cleared", Toast.LENGTH_SHORT).show()
//    }
private fun clearAllData() {
    lifecycleScope.launch {
        clientDao.deleteAll()
        productDao.deleteAll()
        refreshLists()
        selectedClient = null
        selectedProduct = null
       // sectionSelectedClients.clear()
        sectionSelectedProducts.clear()
        tvClientName.text = "Client Name (Dropdown)"
        tvDescriptionMore.text = "Description of Goods (Dropdown)"
        productContainer.removeAllViews()
        Toast.makeText(this@ThirdActivity, "All data cleared", Toast.LENGTH_SHORT).show()
    }
}


    // For dynamic sections (adapt addNewProductSection)
//    private fun addNewProductSection() {
//        val view = layoutInflater.inflate(R.layout.new_product_ly, productContainer, false)
//
//        // Setup dropdowns for new section (similar to setupSectionDropdowns in SecondFragment)
//        setupSectionDropdowns(view)
//
//        // Remove button
//        val removeBtn = view.findViewById<ImageButton>(R.id.btn_remove_product)
//        removeBtn.setOnClickListener {
//            removeProductSection(view)
//        }
//
//        // Quantity listener if needed
//
//        productContainer.addView(view)
//        productSections.add(view)
//        productSectionCounter++
//        Log.d("EXCEL_EXPORT", "NEW SECTION ADDED! Total sections now: ${productSections.size}")
//        Log.d("EXCEL_EXPORT", "productContainer child count: ${productContainer.childCount}")
//        Toast.makeText(this, "New product added", Toast.LENGTH_SHORT).show()
//    }
//    private fun addNewProductSection() {
//        val view = layoutInflater.inflate(R.layout.new_product_ly, productContainer, false)
//
//        setupSectionDropdowns(view)
//
//        val removeBtn = view.findViewById<ImageButton>(R.id.btn_remove_product)
//        removeBtn.setOnClickListener {
//            removeProductSection(view)
//        }
//
//        // ONLY ADD TO productContainer — NEVER touch mainLinearLayout!
//        productContainer.addView(view)
//        productSections.add(view)
//
//        Log.d("EXCEL_EXPORT", "ADDED! Total sections: ${productSections.size} | Children: ${productContainer.childCount}")
//        Toast.makeText(this, "Product ${productSections.size} added", Toast.LENGTH_SHORT).show()
//    }
    private fun addNewProductSection() {
        val view = layoutInflater.inflate(R.layout.new_product_ly, productContainer, false)

        setupSectionDropdowns(view)

        // Correct: Find the REAL + button inside the section
        val btnAddMore = view.findViewById<ImageButton>(R.id.btn_remove_product)
        btnAddMore.setOnClickListener {
            addNewProductSection() // Now the + button adds a new section
        }

        // Correct: Remove button removes only itself
        val btnRemove = view.findViewById<ImageButton>(R.id.btn_remove_product)
        btnRemove.setOnClickListener {
            removeProductSection(view)
        }

        productContainer.addView(view)
        productSections.add(view)

// ENABLE EXPORT BUTTON AS SOON AS FIRST PRODUCT IS ADDED
        if (productSections.size == 1) {
            btn_export.isEnabled = true
            btn_export.text = "Export Data"
        }
        //_____
        Log.d("EXCEL_EXPORT", "SUCCESS! Added section. Total: ${productSections.size}")
        Toast.makeText(this, "Product ${productSections.size} added!", Toast.LENGTH_SHORT).show()
    }

    fun removeProductSection(sectionView: View) {
        // Remove dropdown views
        sectionDropdowns[sectionView]?.let { (clientDropdown, productDropdown) ->
            dropdownContainer.removeView(clientDropdown)
            dropdownContainer.removeView(productDropdown)
        }
        sectionDropdowns.remove(sectionView)
       // sectionSelectedClients.remove(sectionView)
        sectionSelectedProducts.remove(sectionView)

//        mainLinearLayout.removeView(sectionView)
        productContainer.removeView(sectionView)
        productSections.remove(sectionView)

        // Update section titles
        updateSectionTitles()

        // If no products left → disable export button
        if (productSections.isEmpty()) {
            btn_export.isEnabled = false
            btn_export.text = "Export Data"
        }

        Log.d("EXCEL_EXPORT", "Section removed! Now total: ${productSections.size}")
        Toast.makeText(this, "Product section removed!", Toast.LENGTH_SHORT).show()
    }

    private fun setupSectionDropdowns(sectionView: View) {
        val inflater = LayoutInflater.from(this)

        // Create dropdown views for this section
        val clientDropdownView = inflater.inflate(R.layout.dropdown_layout, dropdownContainer, false)
        val productDropdownView = inflater.inflate(R.layout.dropdown_layout, dropdownContainer, false)

        val layoutParams = FrameLayout.LayoutParams(
            FrameLayout.LayoutParams.MATCH_PARENT,
            FrameLayout.LayoutParams.WRAP_CONTENT
        )
        clientDropdownView.layoutParams = layoutParams
        productDropdownView.layoutParams = layoutParams

        clientDropdownView.visibility = View.GONE
        productDropdownView.visibility = View.GONE

        dropdownContainer.addView(clientDropdownView)
        dropdownContainer.addView(productDropdownView)

        // Store dropdown references for this section
        sectionDropdowns[sectionView] = Pair(clientDropdownView, productDropdownView)

        // Setup click listeners for this section
        val clientLayout = sectionView.findViewById<LinearLayout>(R.id.client_name_layout)
        val productLayout = sectionView.findViewById<LinearLayout>(R.id.description_of_goods_layout)
        val clientArrow = sectionView.findViewById<ImageView>(R.id.client_arrow)
        val productArrow = sectionView.findViewById<ImageView>(R.id.product_arrow)

        var isSectionClientDropdownOpen = false
        var isSectionProductDropdownOpen = false

        clientLayout.setOnClickListener {
            toggleSectionClientDropdown(sectionView, clientDropdownView, clientArrow,
                isSectionClientDropdownOpen, isSectionProductDropdownOpen).let { result ->
                isSectionClientDropdownOpen = result.first
                isSectionProductDropdownOpen = result.second
            }
        }

        productLayout.setOnClickListener {
            toggleSectionProductDropdown(sectionView, productDropdownView, productArrow,
                isSectionClientDropdownOpen, isSectionProductDropdownOpen).let { result ->
                isSectionClientDropdownOpen = result.first
                isSectionProductDropdownOpen = result.second
            }
        }

        clientArrow.setOnClickListener {
            toggleSectionClientDropdown(sectionView, clientDropdownView, clientArrow,
                isSectionClientDropdownOpen, isSectionProductDropdownOpen).let { result ->
                isSectionClientDropdownOpen = result.first
                isSectionProductDropdownOpen = result.second
            }
        }

        productArrow.setOnClickListener {
            toggleSectionProductDropdown(sectionView, productDropdownView, productArrow,
                isSectionClientDropdownOpen, isSectionProductDropdownOpen).let { result ->
                isSectionClientDropdownOpen = result.first
                isSectionProductDropdownOpen = result.second
            }
        }
    }


    public fun getFileName(uri: Uri): String {
        return try {
            uri.toString().substringAfterLast("/")
        } catch (e: Exception) {
            "unknown"
        }
    }
    public fun detectFileType(fileName: String): String {
        return when {
            fileName.contains("client", ignoreCase = true) -> "client"
            fileName.contains("product", ignoreCase = true) -> "product"
            fileName.contains("customer", ignoreCase = true) -> "client"
            else -> "auto"
        }
    }
    fun readFileContent(inputStream: InputStream, fileType: String, fileName: String) {
        try {
            // Read all lines first and store them
            val lines = inputStream.bufferedReader().use { it.readLines() }

            when (fileType) {
                "client" -> processClientData(lines, fileName)
                "product" -> processProductData(lines, fileName)
                else -> autoDetectAndProcessData(lines, fileName)
            }

//            updateDisplayItems()
//            updateUI()
        } catch (e: Exception) {
            showMessage("Error processing file: ${e.message}")
        }
    }
    fun processClientData(lines: List<String>, fileName: String) {
        if (lines.isEmpty()) {
            showMessage("File is empty")
            return
        }

        val newClients = mutableListOf<Client>()

        // Check if header exists
        val firstLine = lines[0].lowercase()
        val hasHeader = firstLine.contains("name") || firstLine.contains("client") || firstLine.contains("email")

        val startIndex = if (hasHeader) 1 else 0

        for (i in startIndex until lines.size) {
            try {
                val line = lines[i].trim()
                if (line.isEmpty()) continue

                // Handle different separators: comma, pipe, or tab
                val columns = when {
                    line.contains("|") -> line.split("|").map { it.trim() }
                    line.contains("\t") -> line.split("\t").map { it.trim() }
                    else -> line.split(",").map { it.trim() }
                }

                if (columns[0].isEmpty()) continue

                val client = when (columns.size) {
                    1 -> Client(name = columns[0])
                    2 -> Client(name = columns[0], email = columns[1])
                    3 -> Client(name = columns[0], email = columns[1], phone = columns[2])
                    else -> Client(name = columns[0])
                }
                newClients.add(client)
            } catch (e: Exception) {
                // Skip invalid rows but continue processing
                continue
            }
        }

//        if (newClients.isNotEmpty()) {
//            clientList.addAll(newClients)
//            showMessage("✅ Added ${newClients.size} clients from $fileName")
//        } else {
//            showMessage("No valid client data found in $fileName")
//        }
        //------ for the db room
        if (newClients.isNotEmpty()) {
            lifecycleScope.launch {
                clientDao.deleteAll() // Delete previous
                clientDao.insertAll(*newClients.toTypedArray()) // Save new
                refreshClientList() // Reload list
                showMessage("✅ Added ${newClients.size} clients from $fileName (replaced previous)")
            }
        } else {
            showMessage("No valid client data found in $fileName")
        }
    }
    fun processProductData(lines: List<String>, fileName: String) {
        if (lines.isEmpty()) {
            showMessage("File is empty")
            return
        }

        val newProducts = mutableListOf<ProductItem>()

        // Check if header exists
        val firstLine = lines[0].lowercase()
        val hasHeader = firstLine.contains("product") || firstLine.contains("name") || firstLine.contains("price") || firstLine.contains("description")

        val startIndex = if (hasHeader) 1 else 0

        for (i in startIndex until lines.size) {
            try {
                val line = lines[i].trim()
                if (line.isEmpty()) continue

                // Handle different separators: comma, pipe, or tab
                val columns = when {
                    line.contains("|") -> line.split("|").map { it.trim() }
                    line.contains("\t") -> line.split("\t").map { it.trim() }
                    else -> line.split(",").map { it.trim() }
                }

                if (columns[0].isEmpty()) continue

                val product = when (columns.size) {
                    1 -> ProductItem(name = columns[0])
                    2 -> ProductItem(name = columns[0], description = columns[1])
                    3 -> ProductItem(
                        name = columns[0],
                        description = columns[1],
                        price = columns[2].toDoubleOrNull() ?: 0.0
                    )
                    4 -> ProductItem(
                        name = columns[0],
                        description = columns[1],
                        price = columns[2].toDoubleOrNull() ?: 0.0,
                        quantity = columns[3].toIntOrNull() ?: 0
                    )
                    else -> ProductItem(name = columns[0])
                }
                newProducts.add(product)
            } catch (e: Exception) {
                // Skip invalid rows but continue processing
                continue
            }
        }
        //------ for the db room
        if (newProducts.isNotEmpty()) {
            lifecycleScope.launch {
                productDao.deleteAll() // Delete previous
                productDao.insertAll(*newProducts.toTypedArray()) // Save new
                refreshProductList() // Reload list
                showMessage("✅ Added ${newProducts.size} products from $fileName (replaced previous)")
            }
        } else {
            showMessage("No valid product data found in $fileName")
        }
    }
    fun autoDetectAndProcessData(lines: List<String>, fileName: String) {
        if (lines.isEmpty()) {
            showMessage("File is empty")
            return
        }

        val firstLine = lines[0].lowercase()

        // Simple detection logic based on content
        val isLikelyProduct = lines.any { line ->
            val columns = line.split(",").map { it.trim() }
            columns.size >= 3 && columns[2].toDoubleOrNull() != null
        }

        when {
            firstLine.contains("product") || firstLine.contains("price") || isLikelyProduct -> {
                processProductData(lines, fileName)
            }
            firstLine.contains("client") || firstLine.contains("customer") -> {
                processClientData(lines, fileName)
            }
            else -> {
                // Default to client if we can't determine
                processClientData(lines, fileName)
            }
        }
    }
    fun updateSectionTitles() {
        productSections.forEachIndexed { index, sectionView ->
            val sectionTitle = sectionView.findViewById<TextView>(R.id.et_Client_Name)
            // Only update if it's still the default text
            if (sectionTitle.text.toString().startsWith("Product Section") ||
                sectionTitle.text.toString() == "Select Client") {
                sectionTitle.text = "Product Section ${index + 1}"
            }
        }
        productSectionCounter = productSections.size + 1
    }

    private fun toggleSectionProductDropdown(
        sectionView: View,
        productDropdownView: View,
        productArrow: ImageView,
        isClientOpen: Boolean,
        isProductOpen: Boolean
    ): Pair<Boolean, Boolean> {
        return if (isProductOpen) {
            hideSectionProductDropdown(productDropdownView, productArrow)
            Pair(isClientOpen, false)
        } else {
            showSectionProductDropdown(sectionView, productDropdownView, productArrow)
            // Close client dropdown if open
            if (isClientOpen) {
                val (clientDropdownView, _) = sectionDropdowns[sectionView]!!
//                val clientArrow = sectionView.findViewById<ImageButton>(R.id.client_arrow)
                val clientArrow = sectionView.findViewById<ImageView>(R.id.client_arrow)
                hideSectionClientDropdown(clientDropdownView, clientArrow)
                Pair(false, true)
            } else {
                Pair(false, true)
            }
        }
    }
private fun showSectionClientDropdown(sectionView: View, dropdownView: View, arrow: ImageView) {
    if (clientList.isEmpty()) {
        Toast.makeText(this, "No clients available", Toast.LENGTH_SHORT).show()
        return
    }

    val clientNames = clientList.map { it.name }

    setupDropdownWithSearch( dropdownView, clientNames) { selectedName ->
        val client = clientList.find { it.name == selectedName }
       // sectionSelectedClients[sectionView] = client
        sectionView.findViewById<TextView>(R.id.et_Client_Name).text = selectedName
        hideSectionClientDropdown(dropdownView, arrow)
    }

    val anchor = sectionView.findViewById<LinearLayout>(R.id.client_name_layout)
    showDropdownBelow(dropdownView, anchor)
    rotateArrow(arrow, 180f)
}

    private fun hideSectionClientDropdown(dropdownView: View, arrow: ImageView) {
        dropdownView.visibility = View.GONE
        rotateArrow(arrow, 0f)
    }
    private fun showSectionProductDropdown(sectionView: View, dropdownView: View, arrow: ImageView) {
        if (productList.isEmpty()) {
            Toast.makeText(this, "No products available", Toast.LENGTH_SHORT).show()
            return
        }

        val productNames = productList.map { it.name }
        setupDropdownWithSearch(dropdownView, productNames) { selectedName ->
            val product = productList.find { it.name == selectedName }
            sectionSelectedProducts[sectionView] = product
            sectionView.findViewById<TextView>(R.id.description_text).text = selectedName

            product?.let {
                sectionView.findViewById<EditText>(R.id.et_hsn_code).setText(it.hsnCode ?: "")
                sectionView.findViewById<EditText>(R.id.et_rate) //.setText(it.price.toString())
            }
            hideSectionProductDropdown(dropdownView, arrow)
        }

        showDropdownBelow(dropdownView, sectionView.findViewById(R.id.description_of_goods_layout))
        rotateArrow(arrow, 180f)
    }

    private fun hideSectionProductDropdown(dropdownView: View, arrow: ImageView) {
        dropdownView.visibility = View.GONE
        rotateArrow(arrow, 0f)
    }
    private fun toggleSectionClientDropdown(
        sectionView: View,
        clientDropdownView: View,
        clientArrow: ImageView,
        isClientOpen: Boolean,
        isProductOpen: Boolean
    ): Pair<Boolean, Boolean> {
        return if (isClientOpen) {
            hideSectionClientDropdown(clientDropdownView, clientArrow)
            Pair(false, isProductOpen)
        } else {
            showSectionClientDropdown(sectionView, clientDropdownView, clientArrow)
            // Close product dropdown if open
            if (isProductOpen) {
                val (_, productDropdownView) = sectionDropdowns[sectionView]!!
                val productArrow = sectionView.findViewById<ImageView>(R.id.product_arrow)
                hideSectionProductDropdown(productDropdownView, productArrow)
                Pair(true, false)
            } else {
                Pair(true, false)
            }
        }
    }
    private fun setupSectionDropdownPosition(dropdownView: View, anchorView: View) {
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
            Log.e(SecondFragment.Companion.TAG, "Error setting section dropdown position", e)
        }
    }
    private fun showMessage(message: String) {
//        binding.textviewFirst.text = message
//        binding.textviewFirst.visibility = View.VISIBLE
        Toast.makeText(this, message, Toast.LENGTH_SHORT).show()
    }
    public fun setupDropdownWithSearch(
        dropdownView: View,
        itemList: List<String>,
        onItemSelected: (String) -> Unit
    ) {
        if (itemList.isEmpty()) return

        val recyclerView = dropdownView.findViewById<RecyclerView>(R.id.recyclerViewDropdown)
        val searchView = dropdownView.findViewById<EditText>(R.id.searchView)

        println("setupDropdownWithSearch: Setting up RecyclerView")

        val adapter = DropdownSearchableAdapter(itemList, onItemSelected)
        recyclerView.adapter = adapter
        recyclerView.layoutManager = LinearLayoutManager(this)

        // Add this to ensure RecyclerView is properly sized
        recyclerView.setHasFixedSize(false)

        // Debug: Check RecyclerView visibility and size
        recyclerView.post {
            println("RecyclerView - Width: ${recyclerView.width}, Height: ${recyclerView.height}")
            println("RecyclerView - Visibility: ${recyclerView.visibility}")
            println("RecyclerView - IsShown: ${recyclerView.isShown}")
        }

        // Search filter with EditText
        searchView.addTextChangedListener(object : TextWatcher {
            override fun beforeTextChanged(s: CharSequence?, start: Int, count: Int, after: Int) {}
            override fun onTextChanged(s: CharSequence?, start: Int, before: Int, count: Int) {}
            override fun afterTextChanged(s: Editable?) {
                val query = s.toString()
                println("⌨️check this is  TYPED: '$query'")
                adapter.filter(query)
            }
        })

        searchView.setText("")
        searchView.clearFocus()
    }
    //************************************************Start
    private fun exportToExcelWithPOI(onComplete: () -> Unit) {


        val dynamicProducts = getAllProductData()
        if (dynamicProducts.isEmpty()) {
            Toast.makeText(this, "Please add at least one product", Toast.LENGTH_SHORT).show()
            return
        }

        try {
            val workbook = XSSFWorkbook()
            val sheet = workbook.createSheet("Tax Invoice")

            // === STYLES (make sure these functions exist in your class) ===
            val headerStyle = createHeaderStyle(workbook)
            val titleStyle = createTitleStyle(workbook)
            val boldStyle = createBoldStyle(workbook)
            val normalStyle = createNormalStyle(workbook)
            val borderStyle = createBorderStyle(workbook)
            val centerStyle = createCenterStyle(workbook)
            val dottedStyle = createDottedStyle(workbook)

            // Column widths
            sheet.setColumnWidth(0, 1500)
            sheet.setColumnWidth(1, 2000)
            sheet.setColumnWidth(2, 8000)
            sheet.setColumnWidth(3, 4000)
            sheet.setColumnWidth(4, 2500)
            sheet.setColumnWidth(5, 3000)
            sheet.setColumnWidth(6, 3500)

            var rowIdx = 0  // Only one counter — no more bugs!

            // ==================== HEADER ====================
            val currentDateTime = SimpleDateFormat("dd-MM-yyyy   hh:mm a", Locale.getDefault()).format(Date())

            sheet.createRow(rowIdx++).apply {
                height = 400
                createCell(1).apply { setCellValue("GSTIN : 09ACLPA7672F1ZG"); cellStyle =
                    boldStyle as XSSFCellStyle?
                }
                sheet.addMergedRegion(CellRangeAddress(0, 0, 1, 2))
                createCell(3).apply { setCellValue("TAX INVOICE"); cellStyle =
                    titleStyle as XSSFCellStyle?
                }
                sheet.addMergedRegion(CellRangeAddress(0, 0, 3, 4))
            }

            repeat(4) { sheet.createRow(rowIdx++) }
            sheet.addMergedRegion(CellRangeAddress(1, 4, 1, 6))
            sheet.getRow(2).createCell(1).apply {
                setCellValue("PENTIUM (India)")
                cellStyle = createLogoStyle(workbook) as XSSFCellStyle?
            }

            sheet.createRow(rowIdx++).apply {
                createCell(1).apply {
                    setCellValue("Sales Office: G-17, Gr. Floor, Shiva Tower, G.T. Road, Ghaziabad.")
                    cellStyle = centerStyle as XSSFCellStyle?
                    sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 1, 6))
                }
            }
            sheet.createRow(rowIdx++).apply {
                createCell(1).apply {
                    setCellValue("H.O. : 1-B.M. Complex, 1st Floor, G.T. Road, Ghaziabad.")
                    cellStyle = centerStyle as XSSFCellStyle?
                    sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 1, 6))
                }
            }
            sheet.createRow(rowIdx++).apply {
                createCell(1).apply {
                    setCellValue("M.: 9911177105, 7982739105 - E-mail : pentium.mts@gmail.com")
                    cellStyle = centerStyle as XSSFCellStyle?
                    sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 1, 6))
                }
            }
            addBorderToRegion(sheet, 0, rowIdx - 1, 1, 6, workbook)

            // Order & Date
            sheet.createRow(rowIdx++).apply {
                height = 400
                createCell(1).apply { setCellValue("Order"); cellStyle =
                    boldStyle as XSSFCellStyle?
                }
                createCell(3).apply {
                    setCellValue("Date: ....$invoiceDate  Time: $invoiceDate...................................................................")
                    cellStyle = normalStyle as XSSFCellStyle?
                    sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 3, 6))
                }
            }

            // Client Name
//            val clientName = dynamicProducts.firstOrNull()?.clientName ?: selectedClient?.name ?: "Client"
//            val clientName = if (dynamicProducts.isNotEmpty()) dynamicProducts[0].clientName else "Client"
            val clientName = selectedClient?.name ?: "Client"  // ← Always use global selected client
            sheet.createRow(rowIdx++).apply {
                height = 400
                createCell(1).apply {
                    setCellValue("M/s: ..$clientName.............................................................................")
                    cellStyle = normalStyle as XSSFCellStyle?
                    sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 1, 6))
                }
            }

            // Dotted line
            sheet.createRow(rowIdx++).apply {
                createCell(1).apply {
                    setCellValue(".................................................................................................")
                    cellStyle = dottedStyle as XSSFCellStyle?
                    sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 1, 6))
                }
                addBorderToRegion(sheet, rowIdx - 1, rowIdx - 1, 1, 6, workbook)
            }

            // ==================== TABLE HEADER ====================
            sheet.createRow(rowIdx++).apply {
                height = 500
                listOf("S. No.", "Description of Good", "HSN code", "Qty", "Rate", "Amount")
                    .forEachIndexed { i, text ->
                        createCell(i + 1).apply {
                            setCellValue(text)
                            cellStyle = createTableHeaderStyle(workbook) as XSSFCellStyle?
                        }
                    }
            }

            // ==================== ALL PRODUCTS (FIXED!) ====================
            var subtotal = 0.0
            dynamicProducts.forEachIndexed { index, prod ->
                val amount = prod.quantity * prod.price
                subtotal += amount

                sheet.createRow(rowIdx++).apply {
                    height = 400
                    createCell(1).apply { setCellValue((index + 1).toDouble()); cellStyle =
                        borderStyle as XSSFCellStyle?
                    }
                    createCell(2).apply { setCellValue(prod.name); cellStyle =
                        borderStyle as XSSFCellStyle?
                    }
                    createCell(3).apply { setCellValue(prod.hsnCode ?: ""); cellStyle =
                        borderStyle as XSSFCellStyle?
                    }
                    createCell(4).apply { setCellValue(prod.quantity.toDouble()); cellStyle =
                        borderStyle as XSSFCellStyle?
                    }
                    createCell(5).apply { setCellValue(prod.price); cellStyle =
                        borderStyle as XSSFCellStyle?
                    }
                    createCell(6).apply { setCellValue(amount); cellStyle =
                        borderStyle as XSSFCellStyle?
                    }
                }
            }

            // 5 Empty Rows
            repeat(5) {
                sheet.createRow(rowIdx++).apply {
                    height = 400
                    (1..6).forEach { createCell(it).cellStyle = borderStyle as XSSFCellStyle? }
                }
            }
            addBorderToRegion(sheet, rowIdx - 5, rowIdx - 1, 1, 6, workbook)

            // ==================== TOTALS ====================
            val cgst = cgstRate.text.toString().toDoubleOrNull() ?: 0.0
            val sgst = sgstRate.text.toString().toDoubleOrNull() ?: 0.0
            val igst = igstRate.text.toString().toDoubleOrNull() ?: 0.0

            // Get bank details from EditText fields
            val bankName = etBankName.text.toString().trim()
            val ifscCode = etIfsc.text.toString().trim()
            val accountNo = etAcNo.text.toString().trim()

//            val cgstAmount = subtotal * cgstRate / 100
//            val sgstAmount = subtotal * sgstRate / 100
//            val igstAmount = subtotal * igstRate / 100
            val cgstAmount = subtotal * cgst / 100
            val sgstAmount = subtotal * sgst / 100
            val igstAmount = subtotal * igst / 100

            val grandTotal = subtotal + cgstAmount + sgstAmount + igstAmount

            // Amount in words
            sheet.createRow(rowIdx++).apply {
                createCell(1).apply {
                    setCellValue("Amount in words: ${convertNumberToWords(subtotal.toInt())} Rupees Only")
                    cellStyle = normalStyle as XSSFCellStyle?
                    sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 1, 4))
                }
                createCell(5).apply {
                    setCellValue("Total Amount Before Tax")
                    cellStyle = normalStyle as XSSFCellStyle?
                    sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 5, 6))
                }
                addBorderToRegion(sheet, rowIdx - 1, rowIdx - 1, 1, 6, workbook)
            }

            // CGST, SGST, IGST
            sheet.createRow(rowIdx++).apply { height = 400; createCell(5).apply { setCellValue(" CGST @${cgst}% "); cellStyle =
                normalStyle as XSSFCellStyle?
            }; createCell(6).apply { setCellValue(cgstAmount); cellStyle =
                borderStyle as XSSFCellStyle?
            } }
            sheet.createRow(rowIdx++).apply { height = 400; createCell(5).apply { setCellValue(" SGST @${sgst} % "); cellStyle =
                normalStyle as XSSFCellStyle?
            }; createCell(6).apply { setCellValue(sgstAmount); cellStyle =
                borderStyle as XSSFCellStyle?
            } }
            if (igst > 0) {
                sheet.createRow(rowIdx++).apply { height = 400; createCell(5).apply { setCellValue(" IGST @${igst}% "); cellStyle =
                    normalStyle as XSSFCellStyle?
                }; createCell(6).apply { setCellValue(igstAmount); cellStyle =
                    borderStyle as XSSFCellStyle?
                } }
            }

            // Reverse Charge + Grand Total
            sheet.createRow(rowIdx++).apply {
                height = 400;
                createCell(5).apply { setCellValue("GST on Reverse Charge"); cellStyle =
                normalStyle as XSSFCellStyle?
            };
            sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 5, 6)); addBorderToRegion(sheet, rowIdx - 1, rowIdx - 1, 5, 6, workbook) }

            sheet.createRow(rowIdx++).apply {
                height = 400
//                createCell(1).apply { setCellValue("Bank Name:");
                createCell(1).apply {
                    setCellValue("Bank Name: ${if (bankName.isNotEmpty()) bankName else "..................................."}");
                    cellStyle = normalStyle as XSSFCellStyle?;
                    sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 1, 4)) }

                createCell(5).apply { setCellValue(" GRAND TOTAL: ₹${String.format("%.2f", grandTotal)} ");
                    cellStyle = createGrandTotalStyle(workbook) as XSSFCellStyle?;
                    sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 5, 6)) }
                addBorderToRegion(sheet, rowIdx - 1, rowIdx - 1, 1, 6, workbook)
            }

            // Footer
            sheet.createRow(rowIdx++).apply { createCell(1).apply { setCellValue("Account No: ${if (accountNo.isNotEmpty()) accountNo else "..................................."}"); cellStyle =
                normalStyle as XSSFCellStyle?; sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 1, 4)) }; addBorderToRegion(sheet, rowIdx - 1, rowIdx - 1, 1, 4, workbook) }
            sheet.createRow(rowIdx++).apply { createCell(1).apply {  setCellValue("IFSC Code: ${if (ifscCode.isNotEmpty()) ifscCode else "..................................."}"); cellStyle =
                normalStyle as XSSFCellStyle?; sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 1, 4)) }; addBorderToRegion(sheet, rowIdx - 1, rowIdx - 1, 1, 4, workbook) }
            sheet.createRow(rowIdx++).apply { createCell(1).apply { setCellValue("1. E. & O.E."); cellStyle =
                normalStyle as XSSFCellStyle?; sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 1, 6)) } }
            sheet.createRow(rowIdx++).apply { createCell(1).apply { setCellValue("2. No physical Electrical Damage is covered under warranty."); cellStyle =
                normalStyle as XSSFCellStyle?; sheet.addMergedRegion(CellRangeAddress(rowIdx - 1, rowIdx - 1, 1, 6)) } }

            // ==================== SAVE FILE ====================
            val timestamp = SimpleDateFormat("yyyyMMdd_HHmmss", Locale.getDefault()).format(Date())
            val safeClientName = clientName.replace("[^a-zA-Z0-9]".toRegex(), "_").take(20)
            val fileName = "Invoice_${safeClientName}_$timestamp.xlsx"

            //------------------direct share start
//            val file = File(cacheDir, fileName)
//            FileOutputStream(file).use { outputStream ->
//                workbook.write(outputStream)
//            }
//            workbook.close()
//
//            // Now share the file
//            shareExcelFile(file, fileName)
            // Save file to cache (fast and safe for sharing)
            val file = File(cacheDir, fileName)
            FileOutputStream(file).use { outputStream ->
                workbook.write(outputStream)
            }
            workbook.close()

            // This line shows the share popup (exactly like your screenshot)
            showShareChooser(file, fileName)

            //------------------direct share End

//            if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
//                saveWorkbookToDownloads(workbook, fileName)
//            } else {
//                saveWorkbookToAppDirectory(workbook, fileName)
//            }

            Toast.makeText(this, "Invoice exported successfully!", Toast.LENGTH_LONG).show()


        } catch (e: Exception) {
            Toast.makeText(this, "Export failed: ${e.message}", Toast.LENGTH_LONG).show()
            e.printStackTrace()


        }finally {
            // THIS IS CRITICAL — ALWAYS CALL onComplete()
            onComplete()
        }
    }

    private fun showShareChooser(file: File, fileName: String) {
        val uri = FileProvider.getUriForFile(
            this,
            "${packageName}.provider",  // Keep exactly like this
            file
        )

        val shareIntent = Intent(Intent.ACTION_SEND).apply {
            type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            putExtra(Intent.EXTRA_STREAM, uri)
            putExtra(Intent.EXTRA_SUBJECT, "Tax Invoice - $fileName")
            putExtra(Intent.EXTRA_TEXT, "Please find your invoice attached.\nThank you!\nPentium (India)")
            addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION)
        }

        // This shows the beautiful popup (WhatsApp, Gmail, Instagram, etc.)
        startActivity(Intent.createChooser(shareIntent, "Share Invoice via"))
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
        } + " "
    }
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
    fun getAllProductData(): List<ProductItem> {
        val products = mutableListOf<ProductItem>()
        val globalClientName = selectedClient?.name ?: "Client"

        Log.d("EXCEL_EXPORT", "=== START COLLECTING PRODUCTS ===")
        Log.d("EXCEL_EXPORT", "Total sections in UI: ${productSections.size}")
        Log.d("EXCEL_EXPORT", "Selected Client: $globalClientName")

        productSections.forEachIndexed { index, sectionView ->
            Log.d("EXCEL_EXPORT", "Section $index - View: $sectionView")

            // Get product name from TextView (MOST RELIABLE!)
            val productNameTv = sectionView.findViewById<TextView>(R.id.description_text)
            val productName = productNameTv.text.toString().trim()

            // If no product name → skip this section
            if (productName.isEmpty() || productName == "Description of Goods") {
                Log.d("EXCEL_EXPORT", "Section $index SKIPPED - No product selected")
                return@forEachIndexed
            }

            val hsnEditText = sectionView.findViewById<TextView>(R.id.et_hsn_code)
            val qtyEditText = sectionView.findViewById<EditText>(R.id.et_qty)
            val rateEditText = sectionView.findViewById<EditText>(R.id.et_rate)

            val hsnCode = hsnEditText.text.toString().trim()
            val qty = qtyEditText.text.toString().toIntOrNull() ?: 0
            val rateText = rateEditText.text.toString()
            val rate = if (rateText.isNotEmpty()) rateText.toDouble() else {
                // Fallback: try to get from selected product
                sectionSelectedProducts[sectionView]?.price ?: 0.0
            }

            if (qty <= 0) {
                Log.d("EXCEL_EXPORT", "Section $index SKIPPED - Qty = 0")
                return@forEachIndexed
            }

            products.add(
                ProductItem(
                    hsnCode = hsnCode.ifEmpty { "-" },
                    quantity = qty,
                    name = productName,
                    description = "",
                    price = rate,

                )
            )

            Log.d("EXCEL_EXPORT", "ADDED → $productName | Qty: $qty | Rate: $rate | HSN: $hsnCode")
        }

        Log.d("EXCEL_EXPORT", "TOTAL PRODUCTS ADDED: ${products.size}")
        Log.d("EXCEL_EXPORT", "=== END COLLECTING PRODUCTS ===\n")

        return products
    }
    private fun saveWorkbookToAppDirectory(workbook: Workbook, fileName: String) {
        try {
            val documentsDir =
                File(getExternalFilesDir(Environment.DIRECTORY_DOCUMENTS), "Invoices")
            if (!documentsDir.exists()) documentsDir.mkdirs()

            val file = File(documentsDir, fileName)
            FileOutputStream(file).use { outputStream ->
                workbook.write(outputStream)
                Toast.makeText(this, "Excel saved!\nLocation: $file", Toast.LENGTH_LONG).show()
            }
            workbook.close()
        } catch (e: Exception) {
            Toast.makeText(this, "Failed to save Excel: ${e.message}", Toast.LENGTH_LONG).show()
        }
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
    //************* Room db
    private suspend fun refreshClientList() {
        clientList.clear()
        clientList.addAll(clientDao.getAll())
        updateArrowColors()
    }

    private suspend fun refreshProductList() {
        productList.clear()
        productList.addAll(productDao.getAll())
        updateArrowColors()
    }

    private suspend fun refreshLists() {
        refreshClientList()
        refreshProductList()
    }



}