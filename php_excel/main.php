<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;

if (isset($_POST['export'])) {

    $pdo = new PDO("mysql:host=127.0.0.1;dbname=employee", "root");
    $stmt = $pdo->prepare("SELECT firstname, lastname FROM employee");
    $stmt->execute();

    $spreadsheet = new Spreadsheet();
    $Excel_writer = new Xls($spreadsheet);

    $spreadsheet->setActiveSheetIndex(0);
    $activeSheet = $spreadsheet->getActiveSheet();

    $spreadsheet->getActiveSheet()->mergeCells('A1:B1');
    $activeSheet->setCellValue('A1','Employees')->getStyle('A1')->getFont()->setSize(24)->setBold(true);

    $activeSheet->setCellValue('A2', 'Firstname')->getStyle('A2')->getFont()->setBold(true)->setSize(16)->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);
    $activeSheet->setCellValue('B2', 'Surname')->getStyle('B2')->getFont()->setBold(true)->setSize(16)->getColor()->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);

    $i = 3;
    while ($row = $stmt->fetch()) {
        $activeSheet->setCellValue('A'.$i, $row["firstname"]);
        $activeSheet->setCellValue('B'.$i, $row["lastname"]);
        $i++;
    }

    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment;filename="employees.xls"');
    header('Cache-Control: max-age=0');
    $Excel_writer->save('php://output');

}
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>HTML to PDF</title>

    <link type="text/css" href="style.css" rel="stylesheet"/>
</head>
<body>
<form name="excelform" method="post" action="<?php echo $_SERVER['PHP_SELF']; ?>">
    <button type="submit" name="export" style="margin: 10px">Export to EXCEL</button>
</form>
<table>
    <thead>
    <tr>
        <th>FirstName</th>
        <th>Surname</th>
    </tr>
    </thead>
    <tbody>
    <?php
    $pdo = new PDO("mysql:host=127.0.0.1;dbname=employee", "root");
    $stmt = $pdo->prepare("SELECT firstname, lastname FROM employee");
    $stmt->execute();
    while ($row = $stmt->fetch()) {
        echo "<tr><td>" . $row["firstname"] . '</td><td>' . $row["lastname"] . "</td></tr>";
    }
    ?>
    </tbody>
</table>
</body>
</html>