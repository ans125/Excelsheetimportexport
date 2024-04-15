<?php
// Database configuration
$host = 'localhost';
$dbname = 'phpsheet';
$username = 'root'; // Default username for XAMPP is 'root'
$password = '';     // Default password for XAMPP is empty

// Create connection
$conn = new mysqli($host, $username, $password, $dbname);

// Check connection
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
}

?>
