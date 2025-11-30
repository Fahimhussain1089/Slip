package com.salty.payslip.roomdb

import androidx.room.Dao
import androidx.room.Insert
import androidx.room.OnConflictStrategy
import androidx.room.Query
import com.salty.payslip.model.ProductItem

@Dao
interface ProductDao {
    @Query("SELECT * FROM products")
    suspend fun getAll(): List<ProductItem>

    @Insert(onConflict = OnConflictStrategy.REPLACE) // Replace if name conflict
    suspend fun insertAll(vararg products: ProductItem)

    @Query("DELETE FROM products")
    suspend fun deleteAll()
}