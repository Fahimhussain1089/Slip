package com.salty.payslip.Fragment

import android.os.Bundle
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import androidx.fragment.app.Fragment
import com.salty.payslip.databinding.FragmentFirstBinding
import android.content.Intent
import android.net.Uri
import androidx.activity.result.contract.ActivityResultContracts
import androidx.navigation.fragment.findNavController
import androidx.recyclerview.widget.LinearLayoutManager
import com.salty.payslip.Adapter.Adapter.ProductAdapter
import java.io.BufferedReader
import java.io.InputStreamReader
import  com.salty.payslip.R
import com.salty.payslip.model.Client
import com.salty.payslip.model.ProductItem
//**********different file upload the ***********
class FirstFragment : Fragment() {

    private var _binding: FragmentFirstBinding? = null
    private val binding get() = _binding!!

    private val clientList = mutableListOf<Client>()
    private val productList = mutableListOf<ProductItem>()
    private val displayItems = mutableListOf<Any>()

    private val filePickerLauncher = registerForActivityResult(ActivityResultContracts.StartActivityForResult()) { result ->
        if (result.resultCode == android.app.Activity.RESULT_OK) {
            result.data?.data?.let { uri ->
                readFileFromUri(uri)
            }
        }
    }

    override fun onCreateView(
        inflater: LayoutInflater, container: ViewGroup?,
        savedInstanceState: Bundle?
    ): View {
        _binding = FragmentFirstBinding.inflate(inflater, container, false)
        return binding.root
    }

    override fun onViewCreated(view: View, savedInstanceState: Bundle?) {
        super.onViewCreated(view, savedInstanceState)

        setupRecyclerView()

        binding.buttonUploadClients.text = "Upload Client CSV"
        binding.buttonUploadProducts.text = "Upload Product CSV"
        binding.buttonClearData.text = "Clear All Data"

        binding.buttonUploadClients.setOnClickListener {
            openFilePicker("client")
        }

        binding.buttonUploadProducts.setOnClickListener {
            openFilePicker("product")
        }

        binding.buttonClearData.setOnClickListener {
            clearAllData()
        }

//        binding.buttonFirst.setOnClickListener {
//            findNavController().navigate(R.id.action_FirstFragment_to_SecondFragment)
//        }

        binding.buttonFirst.setOnClickListener {
            // Pass data to SecondFragment using Bundle
            val bundle = Bundle().apply {
                putParcelableArrayList("clientList", ArrayList(clientList))
                putParcelableArrayList("productList", ArrayList(productList))
            }

            findNavController().navigate(R.id.action_FirstFragment_to_SecondFragment, bundle)
        }

    }

        private fun setupRecyclerView() {
        binding.recyclerViewProducts.layoutManager = LinearLayoutManager(requireContext())
        binding.recyclerViewProducts.adapter = ProductAdapter(displayItems)
    }

    private fun openFilePicker(fileType: String) {
        val intent = Intent(Intent.ACTION_GET_CONTENT).apply {
            type = "*/*"
            addCategory(Intent.CATEGORY_OPENABLE)

            val mimeTypes = arrayOf(
                "text/csv",
                "text/comma-separated-values"
            )
            putExtra(Intent.EXTRA_MIME_TYPES, mimeTypes)
        }

        filePickerLauncher.launch(Intent.createChooser(intent, "Select ${fileType.capitalize()} CSV File"))
    }

    private fun readFileFromUri(uri: Uri) {
        try {
            // Get file name to help with detection
            val fileName = getFileName(uri)
            val fileType = detectFileType(fileName)

            requireContext().contentResolver.openInputStream(uri)?.use { inputStream ->
                readFileContent(inputStream, fileType, fileName)
            }
        } catch (e: Exception) {
            e.printStackTrace()
            showMessage("Error reading file: ${e.message}")
        }
    }

    private fun getFileName(uri: Uri): String {
        return try {
            uri.toString().substringAfterLast("/")
        } catch (e: Exception) {
            "unknown"
        }
    }

    private fun detectFileType(fileName: String): String {
        return when {
            fileName.contains("client", ignoreCase = true) -> "client"
            fileName.contains("product", ignoreCase = true) -> "product"
            fileName.contains("customer", ignoreCase = true) -> "client"
            else -> "auto"
        }
    }

    private fun readFileContent(inputStream: java.io.InputStream, fileType: String, fileName: String) {
        try {
            // Read all lines first and store them
            val lines = inputStream.bufferedReader().use { it.readLines() }

            when (fileType) {
                "client" -> processClientData(lines, fileName)
                "product" -> processProductData(lines, fileName)
                else -> autoDetectAndProcessData(lines, fileName)
            }

            updateDisplayItems()
            updateUI()
        } catch (e: Exception) {
            showMessage("Error processing file: ${e.message}")
        }
    }

    private fun processClientData(lines: List<String>, fileName: String) {
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

        if (newClients.isNotEmpty()) {
            clientList.addAll(newClients)
            showMessage("✅ Added ${newClients.size} clients from $fileName")
        } else {
            showMessage("No valid client data found in $fileName")
        }
    }

    private fun processProductData(lines: List<String>, fileName: String) {
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

        if (newProducts.isNotEmpty()) {
            productList.addAll(newProducts)
            showMessage("✅ Added ${newProducts.size} products from $fileName")
        } else {
            showMessage("No valid product data found in $fileName")
        }
    }

    private fun autoDetectAndProcessData(lines: List<String>, fileName: String) {
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

    private fun updateDisplayItems() {
        displayItems.clear()

        if (clientList.isNotEmpty()) {
            displayItems.add("👥 Client List (${clientList.size} clients)")
            displayItems.addAll(clientList)
        }

        if (productList.isNotEmpty()) {
            displayItems.add("📦 Product List (${productList.size} products)")
            displayItems.addAll(productList)
        }
    }

    private fun clearAllData() {
        clientList.clear()
        productList.clear()
        displayItems.clear()
        updateUI()
        showMessage("All data cleared")
    }

    private fun addSampleData() {
        if (clientList.isEmpty() && productList.isEmpty()) {
            clientList.add(Client("comming soon ", "XXXX@email.com", "XXXXXXX210"))

            productList.add(ProductItem("comming soon", "comming soon", 0.0, 0))


            updateDisplayItems()
            updateUI()
            showMessage("Sample data loaded")
        }
    }

    private fun updateUI() {
        binding.recyclerViewProducts.adapter = ProductAdapter(displayItems)

        val totalItems = clientList.size + productList.size
        if (totalItems == 0) {
            showMessage("No data found. Please upload CSV files.")
        } else {
            showMessage("Loaded: ${clientList.size} clients, ${productList.size} products")
        }
    }

    private fun showMessage(message: String) {
        binding.textviewFirst.text = message
        binding.textviewFirst.visibility = View.VISIBLE
    }

    override fun onDestroyView() {
        super.onDestroyView()
        _binding = null
    }
}

//class FirstFragment : Fragment() {
//
//    private var _binding: FragmentFirstBinding? = null
//
//    // This property is only valid between onCreateView and
//    // onDestroyView.
//    private val binding get() = _binding!!
//
//    override fun onCreateView(
//        inflater: LayoutInflater, container: ViewGroup?,
//        savedInstanceState: Bundle?
//    ): View {
//
//        _binding = FragmentFirstBinding.inflate(inflater, container, false)
//        return binding.root
//
//    }
//
//    override fun onViewCreated(view: View, savedInstanceState: Bundle?) {
//        super.onViewCreated(view, savedInstanceState)
//
//        binding.buttonFirst.setOnClickListener {
//            findNavController().navigate(R.id.action_FirstFragment_to_SecondFragment)
//        }
//    }
//
//    override fun onDestroyView() {
//        super.onDestroyView()
//        _binding = null
//    }
//}