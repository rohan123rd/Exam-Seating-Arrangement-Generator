<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER['REQUEST_METHOD'] == 'POST') {
    echo "<pre>";
    print_r($_FILES);
    echo "</pre>";
    
    // Check if files were uploaded
    if (isset($_FILES['masterKeyFile']) && isset($_FILES['seatingArrangementFile'])) {
        $masterKeyFile = $_FILES['masterKeyFile']['tmp_name'];
        $masterKeyFileName = $_FILES['masterKeyFile']['name'];
        $seatingArrangementFile = $_FILES['seatingArrangementFile']['tmp_name'];
        $seatingArrangementFileName = $_FILES['seatingArrangementFile']['name'];

        function loadSpreadsheet($filePath, $fileName) {
            $extension = pathinfo($fileName, PATHINFO_EXTENSION);
            $reader = null;
            
            switch (strtolower($extension)) {
                case 'xlsx':
                    $reader = IOFactory::createReader('Xlsx');
                    break;
                case 'xls':
                    $reader = IOFactory::createReader('Xls');
                    break;
                case 'csv':
                    $reader = IOFactory::createReader('Csv');
                    break;
                default:
                    throw new Exception('Unsupported file type: ' . $extension);
            }
            
            return $reader->load($filePath);
        }

        try {
            $spreadsheetMaster = loadSpreadsheet($masterKeyFile, $masterKeyFileName);
            $spreadsheetSeating = loadSpreadsheet($seatingArrangementFile, $seatingArrangementFileName);

            $sheetMaster = $spreadsheetMaster->getActiveSheet();
            $sheetSeating = $spreadsheetSeating->getActiveSheet();

            $masterData = $sheetMaster->toArray();
            $seatingData = $sheetSeating->toArray();

            // Create an associative array for master data
            $students = [];
            foreach ($masterData as $row) {
                $enrollmentNumber = $row[0]; // Assuming the enrollment number is in the first column
                $studentName = $row[1]; // Assuming the student name is in the second column
                $class = $row[2]; // Assuming the class is in the third column
                $students[$enrollmentNumber] = [
                    'name' => $studentName,
                    'class' => $class
                ];
            }

            // Create the result array
            $result = [];
            foreach ($seatingData as $row) {
                $enrollmentNumber = $row[0]; // Assuming the enrollment number is in the first column
                $subjectCode = $row[1]; // Assuming the subject code is in the second column
                $class = $row[2]; // Assuming the class is in the third column
                if (isset($students[$enrollmentNumber])) {
                    $studentName = $students[$enrollmentNumber]['name'];
                    $result[$class][$subjectCode][] = [
                        'name' => $studentName,
                        'enrollment_number' => $enrollmentNumber
                    ];
                }
            }

            // Create a new Spreadsheet for the result
            $resultSpreadsheet = new Spreadsheet();
            $resultSheet = $resultSpreadsheet->getActiveSheet();
            $row = 1;
            foreach ($result as $class => $subjects) {
                foreach ($subjects as $subjectCode => $studentList) {
                    $resultSheet->setCellValue('A'.$row, "Class: $class");
                    $resultSheet->setCellValue('B'.$row, "Subject: $subjectCode");
                    $row++;
                    $resultSheet->setCellValue('A'.$row, 'Enrollment Number');
                    $resultSheet->setCellValue('B'.$row, 'Student Name');
                    $row++;
                    foreach ($studentList as $student) {
                        $resultSheet->setCellValue('A'.$row, $student['enrollment_number']);
                        $resultSheet->setCellValue('B'.$row, $student['name']);
                        $row++;
                    }
                    $row++;
                }
            }

            // Save the result to a new Excel file in the current directory
            $outputFile = __DIR__ . '/Result_Student_List.xlsx';
            $writer = new Xlsx($resultSpreadsheet);
            $writer->save($outputFile);

            echo "Output file has been generated: <a href='Result_Student_List.xlsx'>Download Result Student List</a>";

        } catch (Exception $e) {
            echo 'Error: ' . $e->getMessage();
        }
    } else {
        echo "No files uploaded.";
    }
} else {
    echo "Invalid request method.";
}
?>