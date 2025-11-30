package com.salty.payslip.roomdb

import androidx.room.Database
import androidx.room.RoomDatabase
import com.salty.payslip.model.Client
import com.salty.payslip.model.ProductItem

@Database(entities = [Client::class, ProductItem::class], version = 1, exportSchema = false)
abstract class AppDatabase : RoomDatabase() {
    abstract fun clientDao(): ClientDao
    abstract fun productDao(): ProductDao
}