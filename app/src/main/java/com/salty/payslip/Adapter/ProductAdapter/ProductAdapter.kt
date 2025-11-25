package com.salty.payslip.Adapter.Adapter

import android.view.LayoutInflater
import android.view.ViewGroup
import android.view.View
import android.widget.TextView
import androidx.recyclerview.widget.RecyclerView
import  com.salty.payslip.R;
import com.salty.payslip.model.Client
import com.salty.payslip.model.ProductItem

// ProductAdapter.kt


//class ProductAdapter(private val items: List<Any>) :
//    RecyclerView.Adapter<ProductAdapter.ItemViewHolder>() {
//
//    companion object {
//        const val TYPE_HEADER = 0
//        const val TYPE_CLIENT = 1
//        const val TYPE_PRODUCT = 2
//    }
//
//    override fun getItemViewType(position: Int): Int {
//        return when (items[position]) {
//            is String -> TYPE_HEADER
//            is Client -> TYPE_CLIENT
//            is ProductItem -> TYPE_PRODUCT
//            else -> TYPE_CLIENT
//        }
//    }
//
//    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): ItemViewHolder {
//        return when (viewType) {
//            TYPE_HEADER -> {
//                val view = LayoutInflater.from(parent.context)
//                    .inflate(R.layout.item_header, parent, false)
//                HeaderViewHolder(view)
//            }
//            TYPE_CLIENT -> {
//                val view = LayoutInflater.from(parent.context)
//                    .inflate(R.layout.item_client, parent, false)
//                ClientViewHolder(view)
//            }
//            else -> {
//                val view = LayoutInflater.from(parent.context)
//                    .inflate(R.layout.item_product, parent, false)
//                ProductViewHolder(view)
//            }
//        }
//    }
//
//    override fun onBindViewHolder(holder: ItemViewHolder, position: Int) {
//        when (holder) {
//            is HeaderViewHolder -> holder.bind(items[position] as String)
//            is ClientViewHolder -> holder.bind(items[position] as Client)
//            is ProductViewHolder -> holder.bind(items[position] as ProductItem)
//        }
//    }
//
//    override fun getItemCount(): Int = items.size
//
//    abstract class ItemViewHolder(itemView: View) : RecyclerView.ViewHolder(itemView)
//
//    class HeaderViewHolder(itemView: View) : ItemViewHolder(itemView) {
//        private val tvHeader: TextView = itemView.findViewById(R.id.tvHeader)
//
//        fun bind(header: String) {
//            tvHeader.text = header
//        }
//    }
//
//    class ClientViewHolder(itemView: View) : ItemViewHolder(itemView) {
//        private val tvClientName: TextView = itemView.findViewById(R.id.tvClientName)
//
//        fun bind(client: Client) {
//            tvClientName.text = client.name
//        }
//    }
//
//    class ProductViewHolder(itemView: View) : ItemViewHolder(itemView) {
//        private val tvProductName: TextView = itemView.findViewById(R.id.tvProductName)
//        private val tvProductDescription: TextView = itemView.findViewById(R.id.tvProductDescription)
//        private val tvProductPrice: TextView = itemView.findViewById(R.id.tvProductPrice)
//
//        fun bind(product: ProductItem) {
//            tvProductName.text = product.name
//            tvProductDescription.text = product.description
//            tvProductPrice.text = "₹${product.price}"
//        }
//    }
//}

//*********different file upload **********


class ProductAdapter(private val items: List<Any>) :
    RecyclerView.Adapter<ProductAdapter.ItemViewHolder>() {

    companion object {
        const val TYPE_HEADER = 0
        const val TYPE_CLIENT = 1
        const val TYPE_PRODUCT = 2
    }

    override fun getItemViewType(position: Int): Int {
        return when (items[position]) {
            is String -> TYPE_HEADER
            is Client -> TYPE_CLIENT
            is ProductItem -> TYPE_PRODUCT
            else -> TYPE_CLIENT
        }
    }

    override fun onCreateViewHolder(parent: ViewGroup, viewType: Int): ItemViewHolder {
        return when (viewType) {
            TYPE_HEADER -> {
                val view = LayoutInflater.from(parent.context)
                    .inflate(R.layout.item_header, parent, false)
                HeaderViewHolder(view)
            }
            TYPE_CLIENT -> {
                val view = LayoutInflater.from(parent.context)
                    .inflate(R.layout.item_client, parent, false)
                ClientViewHolder(view)
            }
            else -> {
                val view = LayoutInflater.from(parent.context)
                    .inflate(R.layout.item_product, parent, false)
                ProductViewHolder(view)
            }
        }
    }

    override fun onBindViewHolder(holder: ItemViewHolder, position: Int) {
        when (holder) {
            is HeaderViewHolder -> holder.bind(items[position] as String)
            is ClientViewHolder -> holder.bind(items[position] as Client)
            is ProductViewHolder -> holder.bind(items[position] as ProductItem)
        }
    }

    override fun getItemCount(): Int = items.size

    abstract class ItemViewHolder(itemView: View) : RecyclerView.ViewHolder(itemView)

    class HeaderViewHolder(itemView: View) : ItemViewHolder(itemView) {
        private val tvHeader: TextView = itemView.findViewById(R.id.tvHeader)

        fun bind(header: String) {
            tvHeader.text = header
        }
    }

    class ClientViewHolder(itemView: View) : ItemViewHolder(itemView) {
        private val tvClientName: TextView = itemView.findViewById(R.id.tvClientName)
        private val tvClientEmail: TextView = itemView.findViewById(R.id.tvClientEmail)
        private val tvClientPhone: TextView = itemView.findViewById(R.id.tvClientPhone)

        fun bind(client: Client) {
            tvClientName.text = client.name
            tvClientEmail.text = client.email.ifEmpty { "No email" }
            tvClientPhone.text = client.phone.ifEmpty { "No phone" }
        }
    }

    class ProductViewHolder(itemView: View) : ItemViewHolder(itemView) {
        private val tvProductName: TextView = itemView.findViewById(R.id.tvProductName)
        private val tvProductDescription: TextView = itemView.findViewById(R.id.tvProductDescription)
        private val tvProductPrice: TextView = itemView.findViewById(R.id.tvProductPrice)
        private val tvProductQuantity: TextView = itemView.findViewById(R.id.tvProductQuantity)

        fun bind(product: ProductItem) {
            tvProductName.text = product.name
            tvProductDescription.text = product.description.ifEmpty { "No description" }
            tvProductPrice.text = "₹ ${product.price}"
            tvProductQuantity.text = "Qty: ${product.quantity}"
        }
    }
}