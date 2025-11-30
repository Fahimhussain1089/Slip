package com.salty.payslip.utils

import android.content.Context
import android.view.View
import android.widget.SearchView
import androidx.recyclerview.widget.LinearLayoutManager
import androidx.recyclerview.widget.RecyclerView
import com.salty.payslip.Adapter.DropdownSearchableAdapter
import  com.salty.payslip.R
//
//public fun setupDropdownWithSearch(
//    dropdownView: View,
//    itemList: List<String>,
//    onItemSelected: (String) -> Unit
//) {
//    if (itemList.isEmpty()) return
//
//    val recyclerView = dropdownView.findViewById<RecyclerView>(R.id.recyclerViewDropdown)
//    val searchView = dropdownView.findViewById<SearchView>(R.id.searchView)
//
//    val adapter = DropdownSearchableAdapter(itemList, onItemSelected)
//    recyclerView.adapter = adapter
//    recyclerView.layoutManager = LinearLayoutManager(this)
//
//    // Search filter
//    searchView.setOnQueryTextListener(object : SearchView.OnQueryTextListener {
//        override fun onQueryTextSubmit(query: String?) = false
//        override fun onQueryTextChange(newText: String?): Boolean {
//            adapter.filter(newText.orEmpty())
//            return true
//        }
//    })
//
//    // Optional: Clear search when closed
//    searchView.setQuery("", false)
//    searchView.clearFocus()
//}
//fun setupDropdownWithSearch(
//    context: Context,
//    dropdownView: View,
//    itemList: List<String>,
//    onItemSelected: (String) -> Unit
//) {
//    if (itemList.isEmpty()) return
//
//    val recyclerView = dropdownView.findViewById<RecyclerView>(R.id.recyclerViewDropdown)
//    val searchView = dropdownView.findViewById<SearchView>(R.id.searchView)
//
//    val adapter = DropdownSearchableAdapter(itemList, onItemSelected)
//    recyclerView.adapter = adapter
//    recyclerView.layoutManager = LinearLayoutManager(context)
//
//    searchView.setOnQueryTextListener(object : SearchView.OnQueryTextListener {
//        override fun onQueryTextSubmit(query: String?) = false
//        override fun onQueryTextChange(newText: String?): Boolean {
//            adapter.filter(newText.orEmpty())
//            return true
//        }
//    })
//
//    searchView.setQuery("", false)
//    searchView.clearFocus()
//}
