/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package config;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class DatabaseConnection {
    private static Connection conn;

    public static Connection getConnection() {
        try {
            if (conn == null) {
                String url = "jdbc:mysql://localhost:3306/java_penjadwalanSiswa";
                String user = "root"; // ubah jika kamu pakai user lain
                String pass = "";     // ubah jika kamu punya password MySQL

                Class.forName("com.mysql.jdbc.Driver");
                conn = DriverManager.getConnection(url, user, pass);
                System.out.println("Koneksi berhasil");
            }
        } catch (Exception e) {
            System.err.println("Koneksi gagal: " + e.getMessage());
        }

        return conn;
    }
}

