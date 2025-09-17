
<?php require_once("./config.php");?>

<?php  
    require_once('./vendor/autoload.php');
    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    ini_set('memory_limit', '1024M');
?>
<html>
    <style>
        .box_contain
        {
            background-color: rgba(255, 255, 255, 0.5); 
            padding-top: 20px;
            padding-bottom: 20px;
        }
        .f_box
        {
            background-color: gainsboro;
            width: 400px;
            border: solid;
            margin: auto;
        }
        .f_text
        {
            margin-left: 5px;
            display: table;
        }
        .row
        {
            display: table-row;
        }
        input
        {
            display: table-cell;
            margin-left: 10px;
        }
        label
        {
            display: table-cell;
        }
        body
        {
            width: 100%;
            height: 100%;
            background: skyblue;
        }
        .inst
        {
            background-color: rgba(255, 255, 255, 0.5); 
        }
    </style>
    <head>
        <title>Excel Filter</title>
        <h2>
            Excel Filter
        </h2>
        <p class="inst">
            Selecciona un archivo xlsx. "Remover datos de cálculos:" Selecciona cuales datos no van a ser utilizados en la suma del total. "Mostrar solo estos datos:" Solo los datos que contienen las palabras escritas aquí exactamente, aparecerán en el archivo mostrando la cantidad y el porciento. Escribir coma para seleccionar múltiples palabras. Dejar en blanco para no seleccionar nada y mostrar todo. "Remover columnas:" Escribe letras correspondientes a las columnas de excel para eliminarlas. Ejemplo, A-C Elimina las columnas A,B,C. A,C-D elimina columnas A,C,D.
        </p>
    </head>
    <br><br>
    <body>
        <div class="box_contain">
            <div class="f_box">
                <div class="f_text">
                    <?php
                    if(isset($_POST["sent"]))
                    {
                            $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
                            $reader->setReadDataOnly(true);
                            $spreadsheet = $reader->load($_FILES['tables']['tmp_name']);
                   
                            $highestRow = $spreadsheet->getActiveSheet()->getHighestRow();
                            $highestColumn = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($spreadsheet->getActiveSheet()->getHighestDataColumn());
                            $values = [];
                            $x = "A";
                            for($i = 1;$i<=$highestColumn;$i++)
                            {
                               
                                $values[] = $x;
                                $x++;
                            }
                    
                            $ws = $spreadsheet->getActiveSheet();

                            $data = [];
                            
                            if(is_array($values))
                            {
                                foreach ($values as $key)
                                {
                                    $fColumn = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString("{$key}"); 
                                    for($row = 2; $row <= $highestRow ; ++$row)
                                    {
                                        $title = $ws->getCellByColumnAndRow($fColumn, 1)->getValue();
                                        $data["{$title}"][] = $ws->getCellByColumnAndRow($fColumn, $row)->getValue();
                                    }
                                }
                            }
                            else
                            {
                                $values = strtolower($values);
                                $fColumn = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString("$values"); 
                                for($row = 2; $row <= $highestRow ; ++$row)
                                {
                                    $title = $ws->getCellByColumnAndRow($fColumn, 1)->getValue();
                     
                                    $data[$title][] = $ws->getCellByColumnAndRow($fColumn, $row)->getValue();
                                }
                            }
                            
                            
                            $_SESSION["ans"] = $data;
                            
                            $form = <<<DELIMETER
                            <form class="formed" action="./index.php" id="index" method="post">
                            <br>
                            <r class="row">
                                <label>Remover datos de cálculos:</label> <input type="text" name="y" placeholder="No aplica"><br><br>
                                
                            </r>
                            <r class="row">
                                <label>Mostrar solo estos datos:</label> <input type="text" name="x" placeholder="Satisfecho, Muy satisfecho"><br><br>
                                
                            </r>
                            <r class="row">
                                <label>Remover columnas:</label> <input type="text" name="z" placeholder="A-C, S"><br><br>
                                
                            </r>
                            <r class="row">
                                 <input name="remove" type="submit"><br>
                            </r>
                        DELIMETER;
                            echo $form;
                            $form = <<<DELIMETER
                                    </form>
                                DELIMETER;
                            echo $form;

                    }
     
            
                        
                    else if(empty($_POST["remove"]))
                    {
                        $form = <<<DELIMETER
                        <form name="form1" id="form1" action="./index.php" method="post" enctype="multipart/form-data">
                            <br>
                            <r class="row">
                                <label>Archivo xlsx:</label> <input type="file" name="tables" accept=".xls,.xlsx, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel" required /> <br><br>
                            </r>
                      
      
                            <input name="sent" type="submit">
                        </form>
                        DELIMETER;
                        echo $form;
                    }
                    ?>
                    <?php
                    $success;
                    if(!empty($_POST["remove"]))
                    {
                        if((empty($_POST["x"]) && empty($_POST["y"])) && empty($_POST["z"]))
                        {
                            $success = filter($_SESSION["ans"], null, null, null);
                        }
                        else if (empty($_POST["x"]) && empty($_POST["z"] ))
                        {
                            $success = filter($_SESSION["ans"], null, $_POST["y"],null);
                        }
                        else if (empty($_POST["y"])&& empty($_POST["z"] ))
                        {
                            $success = filter($_SESSION["ans"], $_POST["x"], null,null);
                        }
                        else if (empty($_POST["y"])&& empty($_POST["x"] ))
                        {
                            $success = filter($_SESSION["ans"], null, null, $_POST["z"]);
                        }
                        else if (empty($_POST["x"] ))
                        {
                            $success = filter($_SESSION["ans"], null, $_POST["y"], $_POST["z"]);
                        }
                        else if (empty($_POST["y"]))
                        {
                            $success = filter($_SESSION["ans"], $_POST["x"], null, $_POST["z"]);
                        }
                        else if (empty($_POST["z"]))
                        {
                            $success = filter($_SESSION["ans"], $_POST["x"], $_POST["y"], null);
                        }
                        else 
                        {
                            $success = filter($_SESSION["ans"], $_POST["x"], $_POST["y"]);
                        }
                        if($success)
                        {
                            redirect("index.php");
                        }
              
                    }

                    ?>
                </div>
            </div>
        </div>
    </body>
</html>