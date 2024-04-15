<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <title>Import and Export Excel using PHP OOP MySQL</title>
    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <style>
        body {
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            background-color: #f4f4f4;
        }

        .container {
            background-color: #ffffff;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
        }

        .btn {
            margin-top: 10px;
        }

        .message {
            margin-top: 20px;
            padding: 10px;
            border-radius: 5px;
            text-align: center;
        }
    </style>
</head>

<body>

    <div class="container">
        <h2 >Import Excel Data into MySQL</h2>

        <form action="code.php" method="post" enctype="multipart/form-data">
            <div class="form-group">
                <label>Select Excel file to upload:</label>
                <input type="file" name="excel_file" class="form-control-file" required>
            </div>
            <input type="submit" value="Import" name="import" class="btn btn-primary">
        </form>

        <?php if (isset($_SESSION['import_message'])) : ?>
            <div class="message alert alert-success">
                <?php echo $_SESSION['import_message']; ?>
            </div>
        <?php unset($_SESSION['import_message']); endif; ?>

        <br>

        <h2 >Export MySQL Data to Excel</h2>

        <form action="code.php" method="post">
            <input type="submit" value="Export" name="export" class="btn btn-success">
        </form>
    </div>

    <!-- Bootstrap JS and Popper.js -->
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/popper.js/1.16.0/umd/popper.min.js"></script>
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.5.2/js/bootstrap.min.js"></script>

</body>

</html>
