<?php

class ExcelHandler {
    private $conn;

    public function __construct($conn) {
        $this->conn = $conn;
    }

    public function convertAndImport() {
        if (isset($_POST['import'])) {
            if ($_FILES['excel_file']['name']) {
                // Check file format
                $file_ext = pathinfo($_FILES['excel_file']['name'], PATHINFO_EXTENSION);
                if (!in_array($file_ext, ['xlsx', 'xls', 'csv'])) {
                    echo '<div class="alert alert-danger text-center ">Unsupported file format!</div>';
                    return;
                }

                $filename = $_FILES['excel_file']['tmp_name'];

                // Check for null file
                if ($_FILES['excel_file']['size'] == 0) {
                    echo '<div class="alert alert-danger text-center">Please select a file!</div>';
                    return;
                }

                // Remove existing data from the table
                $this->conn->query("TRUNCATE TABLE students");

                // Import Excel data into the database
                $rowCount = $this->importExcelToDatabase($filename, $file_ext);
                
                if ($rowCount !== false) {
                    // Display success message
                    echo '<div class="alert alert-success text-center">Data imported successfully! Rows imported: ' . $rowCount . '</div>';
                }
            }
        }

        if (isset($_POST['export'])) {
            // Fetch data from MySQL
            $sql = "SELECT fullname, email, phone, course FROM students";
            $result = $this->conn->query($sql);

            // Set headers for CSV file download
            header('Content-Type: text/csv');
            header('Content-Disposition: attachment; filename="export_' . date("Ymd_His") . '.csv"');

            // Create CSV file
            $output = fopen('php://output', 'w');

            // Write headers
            fputcsv($output, ['Full Name', 'Email', 'Phone', 'Course']);

            // Write data
            while ($row = $result->fetch_assoc()) {
                fputcsv($output, $row);
            }

            fclose($output);
        }
    }

    private function importExcelToDatabase($filename, $file_ext) {
        $rowCount = 0;

        if ($file_ext == 'xlsx') {
            $xml_content = file_get_contents("zip://$filename#xl/sharedStrings.xml");
            $shared_strings = new SimpleXMLElement($xml_content);
            
            $xml_content = file_get_contents("zip://$filename#xl/worksheets/sheet1.xml");
            $sheet_data = new SimpleXMLElement($xml_content);

            foreach ($sheet_data->sheetData->row as $row) {
                $rowData = [];
                foreach ($row->c as $cell) {
                    $v = (string) $cell->v;
                    if (isset($cell['t']) && (string) $cell['t'] == 's') {
                        $v = (string) $shared_strings->si[(int) $v]->t;
                    }
                    $rowData[] = $v;
                }

                if (count($rowData) != 4) {
                    echo '<div class="alert alert-danger text-center">Error: Incorrect number of fields in row ' . ($rowCount + 1) . '!</div>';
                    return false;
                }

                $fullname = mysqli_real_escape_string($this->conn, $rowData[0]);
                $email = mysqli_real_escape_string($this->conn, $rowData[1]);
                $phone = mysqli_real_escape_string($this->conn, $rowData[2]);
                $course = mysqli_real_escape_string($this->conn, $rowData[3]);

                // Insert data into the database
                $sql = "INSERT INTO students (fullname, email, phone, course) VALUES ('$fullname', '$email', '$phone', '$course')";

                if ($this->conn->query($sql) !== TRUE) {
                    echo '<div class="alert alert-danger text-center">Error importing data: ' . $this->conn->error . '</div>';
                    return false;
                }
                
                $rowCount++;
            }

            return $rowCount;
        } elseif ($file_ext == 'xls') {
            echo '<div class="alert alert-danger text-center">XLS format is not supported directly. Please convert it to XLSX or CSV format.</div>';
            return false;
        } elseif ($file_ext == 'csv') {
            if (($handle = fopen($filename, "r")) !== FALSE) {
                while (($data = fgetcsv($handle, 1000, ",")) !== FALSE) {
                    $rowCount++;

                    if (count($data) != 4) {
                        echo '<div class="alert alert-danger text-center">Error: Incorrect number of fields in row ' . $rowCount . '!</div>';
                        fclose($handle);
                        return false;
                    }

                    $fullname = mysqli_real_escape_string($this->conn, $data[0]);
                    $email = mysqli_real_escape_string($this->conn, $data[1]);
                    $phone = mysqli_real_escape_string($this->conn, $data[2]);
                    $course = mysqli_real_escape_string($this->conn, $data[3]);

                    // Insert data into the database
                    $sql = "INSERT INTO students (fullname, email, phone, course) VALUES ('$fullname', '$email', '$phone', '$course')";

                    if ($this->conn->query($sql) !== TRUE) {
                        echo '<div class="alert alert-danger text-center">Error importing data: ' . $this->conn->error . '</div>';
                        fclose($handle);
                        return false;
                    }
                }
                fclose($handle);
                return $rowCount;
            } else {
                echo '<div class="alert alert-danger text-center">Error reading CSV file!</div>';
                return false;
            }
        } else {
            echo '<div class="alert alert-danger text-center">Unsupported file format!</div>';
            return false;
        }
    }
}

// Include DB connection
include 'config.php';

$handler = new ExcelHandler($conn);

$handler->convertAndImport();

?>
