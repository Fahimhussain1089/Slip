package com.salty.payslip.Adapter

import android.R.attr.textSize
import android.graphics.Color
import android.view.LayoutInflater
import android.view.View
import android.view.ViewGroup
import android.widget.TextView
import androidx.recyclerview.widget.RecyclerView

class DropdownSearchableAdapter(
    private val originalList: List<String>,
    private val onItemClick: (String) -> Unit
) : RecyclerView.Adapter<DropdownSearchableAdapter.ViewHolder>() {

    private var filteredList: List<String> = originalList

    inner class ViewHolder(itemView: View) : RecyclerView.ViewHolder(itemView) {
        val textView: TextView = itemView.findViewById(android.R.id.text1)

        init {
            // Make the item clickable
            itemView.setOnClickListener {
                val position = adapterPosition
                if (position != RecyclerView.NO_POSITION) {
                    val item = filteredList[position]
                    println("check this is Item clicked: $item at position $position")
                    onItemClick(item)
                }
            }
        }
    }

    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): ViewHolder {
        val view = LayoutInflater.from(parent.context)
            .inflate(android.R.layout.simple_list_item_1, parent, false)

        val textView = view.findViewById<TextView>(android.R.id.text1)
        textView.setPadding(40, 32, 40, 32)
        textView.textSize = 15f
        textView.setTextColor(Color.BLACK) // Ensure text is visible

        return ViewHolder(view)
    }

    override fun onBindViewHolder(holder: ViewHolder, position: Int) {
        val item = filteredList[position]
        holder.textView.text = item
        println("check this is Binding item: $item at position $position")
    }

    override fun getItemCount(): Int {
        val count = filteredList.size
        println("check this is getItemCount called: $count")
        return count
    }

    fun filter(query: String) {
        filteredList = if (query.isEmpty()) {
            originalList
        } else {
            originalList.filter {
                it.contains(query, ignoreCase = true)
            }
        }

        // Simplified debug logs
        println("🔍check this is FILTER: Query='$query' | Results=${filteredList.size}")

        notifyDataSetChanged()
    }
}