package com.salty.payslip.Fragment

import android.os.Bundle
import android.util.Log
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.FrameLayout
import android.widget.ImageButton
import android.widget.TextView
import android.widget.Toast
import androidx.coordinatorlayout.widget.CoordinatorLayout
import androidx.fragment.app.Fragment
import androidx.navigation.fragment.findNavController
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import com.salty.payslip.Adapter.DropdownAdapter
import com.salty.payslip.R
import com.salty.payslip.databinding.FragmentSecondBinding
import com.salty.payslip.model.Client
import com.salty.payslip.model.ProductItem

class SecondFragment : Fragment() {

    private var _binding: FragmentSecondBinding? = null
    private val binding get() = _binding!!

    private val clientList = mutableListOf<Client>()
    private val productList = mutableListOf<ProductItem>()

    private var isClientDropdownOpen = false
    private var isProductDropdownOpen = false

    private lateinit var clientDropdownView: View
    private lateinit var productDropdownView: View

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

        // Log all data for debugging
        logAllData()

        setupDropdowns()
        setupClickListeners()

        binding.buttonSecond.setOnClickListener {
            findNavController().navigate(R.id.action_SecondFragment_to_FirstFragment)
        }
    }

    private fun logAllData() {
        // Log basic counts
        Log.d("SecondFragment", "=== DATA RECEIVED FROM FIRST FRAGMENT ===")
        Log.d("SecondFragment", "Total clients: ${clientList.size}")
        Log.d("SecondFragment", "Total products: ${productList.size}")

        // Log all client details
        if (clientList.isEmpty()) {
            Log.d("SecondFragment", "CLIENT LIST: EMPTY")
        } else {
            Log.d("SecondFragment", "=== CLIENT LIST ===")
            clientList.forEachIndexed { index, client ->
                Log.d("SecondFragment", "Client ${index + 1}:")
                Log.d("SecondFragment", "  - Name: ${client.name}")
                Log.d("SecondFragment", "  - Email: ${client.email}")
                Log.d("SecondFragment", "  - Phone: ${client.phone}")
            }
        }

        // Log all product details
        if (productList.isEmpty()) {
            Log.d("SecondFragment", "PRODUCT LIST: EMPTY")
        } else {
            Log.d("SecondFragment", "=== PRODUCT LIST ===")
            productList.forEachIndexed { index, product ->
                Log.d("SecondFragment", "Product ${index + 1}:")
                Log.d("SecondFragment", "  - Name: ${product.name}")
                Log.d("SecondFragment", "  - Description: ${product.description}")
                Log.d("SecondFragment", "  - Price: ${product.price}")
                Log.d("SecondFragment", "  - Quantity: ${product.quantity}")
            }
        }

        // Log just the names for dropdowns
        val clientNames = clientList.map { it.name }
        val productNames = productList.map { it.name }

        Log.d("SecondFragment", "=== DROPDOWN DATA ===")
        Log.d("SecondFragment", "Client Names: $clientNames")
        Log.d("SecondFragment", "Product Names: $productNames")
        Log.d("SecondFragment", "=== END DATA LOG ===")
    }

    private fun setupDropdowns() {
        // Create dropdown views with proper layout parameters
        val inflater = LayoutInflater.from(requireContext())

        // Inflate without attaching to parent initially
        clientDropdownView = inflater.inflate(R.layout.dropdown_layout, binding.dropdownContainer, false)
        productDropdownView = inflater.inflate(R.layout.dropdown_layout, binding.dropdownContainer, false)

        // Set proper layout parameters for FrameLayout
        val layoutParams = FrameLayout.LayoutParams(
            FrameLayout.LayoutParams.MATCH_PARENT,
            FrameLayout.LayoutParams.WRAP_CONTENT
        )
        clientDropdownView.layoutParams = layoutParams
        productDropdownView.layoutParams = layoutParams

        // Initially hide dropdowns
        clientDropdownView.visibility = View.GONE
        productDropdownView.visibility = View.GONE

        // Add to dropdown container
        binding.dropdownContainer.addView(clientDropdownView)
        binding.dropdownContainer.addView(productDropdownView)

        // Log dropdown setup
        Log.d("SecondFragment", "Dropdowns setup completed")
    }

    private fun setupClickListeners() {
        // Client Name dropdown
        binding.clientNameLayout.setOnClickListener {
            Log.d("SecondFragment", "Client Name layout clicked")
            toggleClientDropdown()
        }

        // Description of Goods dropdown
        binding.descriptionOfGoodsLayout.setOnClickListener {
            Log.d("SecondFragment", "Description layout clicked")
            toggleProductDropdown()
        }

        // Arrow buttons
        binding.clientArrow.setOnClickListener {
            Log.d("SecondFragment", "Client arrow clicked")
            toggleClientDropdown()
        }

        binding.productArrow.setOnClickListener {
            Log.d("SecondFragment", "Product arrow clicked")
            toggleProductDropdown()
        }

        Log.d("SecondFragment", "Click listeners setup completed")
    }

    private fun toggleClientDropdown() {
        Log.d("SecondFragment", "toggleClientDropdown() - Current state: $isClientDropdownOpen")
        if (isClientDropdownOpen) {
            hideClientDropdown()
        } else {
            showClientDropdown()
            hideProductDropdown()
        }
    }

    private fun toggleProductDropdown() {
        Log.d("SecondFragment", "toggleProductDropdown() - Current state: $isProductDropdownOpen")
        if (isProductDropdownOpen) {
            hideProductDropdown()
        } else {
            showProductDropdown()
            hideClientDropdown()
        }
    }

    private fun showClientDropdown() {
        Log.d("SecondFragment", "showClientDropdown() - Client list size: ${clientList.size}")

        if (clientList.isEmpty()) {
            Log.w("SecondFragment", "No clients available to show in dropdown")
            Toast.makeText(requireContext(), "No clients available", Toast.LENGTH_SHORT).show()
            return
        }

        val clientNames = clientList.map { it.name }
        Log.d("SecondFragment", "Client names for dropdown: $clientNames")

        val recyclerView = clientDropdownView.findViewById<RecyclerView>(R.id.recyclerViewDropdown)
        recyclerView.layoutManager = LinearLayoutManager(requireContext())
        recyclerView.adapter = DropdownAdapter(clientNames) { selectedClient ->
            Log.d("SecondFragment", "Client selected: $selectedClient")
            binding.ClientName.text = selectedClient
            hideClientDropdown()
        }

        // Position the dropdown below Client Name
        setupDropdownPosition(clientDropdownView, binding.clientNameLayout)

        clientDropdownView.visibility = View.VISIBLE
        isClientDropdownOpen = true

        rotateArrow(binding.clientArrow, 180f)
        Log.d("SecondFragment", "Client dropdown shown successfully")
    }

    private fun hideClientDropdown() {
        Log.d("SecondFragment", "hideClientDropdown()")
        clientDropdownView.visibility = View.GONE
        isClientDropdownOpen = false
        rotateArrow(binding.clientArrow, 0f)
    }

    private fun showProductDropdown() {
        Log.d("SecondFragment", "showProductDropdown() - Product list size: ${productList.size}")

        if (productList.isEmpty()) {
            Log.w("SecondFragment", "No products available to show in dropdown")
            Toast.makeText(requireContext(), "No products available", Toast.LENGTH_SHORT).show()
            return
        }

        val productNames = productList.map { it.name }
        Log.d("SecondFragment", "Product names for dropdown: $productNames")

        val recyclerView = productDropdownView.findViewById<RecyclerView>(R.id.recyclerViewDropdown)
        recyclerView.layoutManager = LinearLayoutManager(requireContext())
        recyclerView.adapter = DropdownAdapter(productNames) { selectedProduct ->
            Log.d("SecondFragment", "Product selected: $selectedProduct")
            binding.descriptionText.text = selectedProduct
            hideProductDropdown()
        }

        // Position the dropdown below Description
        setupDropdownPosition(productDropdownView, binding.descriptionOfGoodsLayout)

        productDropdownView.visibility = View.VISIBLE
        isProductDropdownOpen = true

        rotateArrow(binding.productArrow, 180f)
        Log.d("SecondFragment", "Product dropdown shown successfully")
    }

    private fun hideProductDropdown() {
        Log.d("SecondFragment", "hideProductDropdown()")
        productDropdownView.visibility = View.GONE
        isProductDropdownOpen = false
        rotateArrow(binding.productArrow, 0f)
    }

    private fun setupDropdownPosition(dropdownView: View, anchorView: View) {
        try {
            val location = IntArray(2)
            anchorView.getLocationOnScreen(location)

            // Get current layout params or create new ones - use FrameLayout.LayoutParams
            val params = dropdownView.layoutParams as? FrameLayout.LayoutParams
                ?: FrameLayout.LayoutParams(
                    FrameLayout.LayoutParams.MATCH_PARENT,
                    FrameLayout.LayoutParams.WRAP_CONTENT
                )

            params.topMargin = location[1] + anchorView.height - getStatusBarHeight()
            dropdownView.layoutParams = params

            Log.d("SecondFragment", "Dropdown positioned - Top margin: ${params.topMargin}")
        } catch (e: Exception) {
            Log.e("SecondFragment", "Error setting dropdown position", e)
        }
    }

    private fun getStatusBarHeight(): Int {
        var result = 0
        val resourceId = resources.getIdentifier("status_bar_height", "dimen", "android")
        if (resourceId > 0) {
            result = resources.getDimensionPixelSize(resourceId)
        }
        Log.d("SecondFragment", "Status bar height: $result")
        return result
    }

    private fun rotateArrow(imageButton: ImageButton, degrees: Float) {
        Log.d("SecondFragment", "Rotating arrow to $degrees degrees")
        imageButton.animate().rotation(degrees).setDuration(300).start()
    }

    override fun onDestroyView() {
        Log.d("SecondFragment", "onDestroyView()")
        super.onDestroyView()
        _binding = null
    }
}