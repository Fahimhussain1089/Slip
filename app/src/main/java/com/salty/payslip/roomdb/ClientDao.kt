package com.salty.payslip.roomdb

import androidx.room.Dao
import androidx.room.Insert
import androidx.room.OnConflictStrategy
import androidx.room.Query
import com.salty.payslip.model.Client

@Dao
interface ClientDao {
    @Query("SELECT * FROM clients")
    suspend fun getAll(): List<Client>

    @Insert(onConflict = OnConflictStrategy.REPLACE) // Replace if name conflict
    suspend fun insertAll(vararg clients: Client)

    @Query("DELETE FROM clients")
    suspend fun deleteAll()
}