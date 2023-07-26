<?php

    $name = $age = $email = "";

    
        // Retrieve form data
        if(isset($_POST['name']));
        $age = $_POST['age'];
        $email = $_POST['email'];

        // Create a new Word document using a template
        $templatePath = 'C:\Users\email\Desktop\Hello.docx';
        $outputPath = 'C:\Users\email\Desktop\Letter.docx';
        copy($templatePath, $outputPath); // Copy the template to a new file

        // Open the Word document
        $document = new COM("Word.Application");
        $document->Documents->Open($outputPath);

        // Replace placeholders in the template with form data
        $document->ActiveDocument->Content->Find->Execute("{{name}}", false, false, false, false, false, true, 1, true, $name);
        $document->ActiveDocument->Content->Find->Execute("{{age}}", false, false, false, false, false, true, 1, true, $age);
        $document->ActiveDocument->Content->Find->Execute("{{email}}", false, false, false, false, false, true, 1, true, $email);

        // Save and close the Word document
        $document->ActiveDocument->Save();
        $document->Quit();

        // Display a success message
        echo "Form data saved to Word document successfully!";
    ?>

