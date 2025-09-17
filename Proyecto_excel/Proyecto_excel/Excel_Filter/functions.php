<?php

$upload_directory = "uploads";






function redirect($location)
{
    return header('refresh:3; url='.$location);
}




function Filter($answers2, $check_quest, $check_answ, $check_col){
    //Excel
    $quest_null = false;
    $answer_null = false;
    $col_null = false;
    if ( is_null($check_quest))
    {
        $check_quest = "";
        $quest_null = true;
    }
    if (is_null($check_answ))
    {
        $check_answ = "";
        $answer_null = true;
    }
    if (is_null($check_col))
    {
        $check_col = "";
        $col_null = true;
    }
    $answers = multi_unique($answers2);
    
    $g = 0;
    $z = 0;
    $answers3 = null;
    $check_answ = preg_replace('/\s*,\s*/', ',', $check_answ);
    $check_quest = preg_replace('/\s*,\s*/', ',', $check_quest);
    $check_col = preg_replace('/\s*,\s*/', ',', $check_col);
    $check_col = str_replace(' ', '', $check_col);
    $check_answ = explode (",",$check_answ);
    $check_quest = explode (",",$check_quest);
    $check_col = explode (",",$check_col);
    $e_columns = [];
    if (is_array ($check_col) && !$col_null)
    {
        foreach($check_col as $t => $value)
        {
            if(str_contains($value,"-"))
            {
                $split = explode("-",$value);
                if(is_array($split))
                {
                    $valid = TRUE;
                    foreach($split as $key)
                    {
                        if(!ctype_alpha($key))
                        {
                            $valid = FALSE;
                        }

                    }
                    if ($valid)
                    {

                        $u_num = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($split[0]);
                        $l_num = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($split[1]);
                        if($u_num > $l_num)
                        {
                            for( $i = $l_num; $i <= $u_num; $i++)
                            {
                                $e_columns[]= $i;
                            }
                        }
                        else
                        {
                            for($i = $u_num; $i <= $l_num; $i++)
                            {
                                $e_columns[]= $i;
                            }
                        }

                    }
                }
                else
                {
                    $col_null = TRUE;
                }
            }
            else
            {
                if(ctype_alpha($value))
                {
                    $e_columns[] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($value);
                }
                else
                {
                    $col_null = TRUE;
                }
            }
        }
    }
    else
    {
        if(!$col_null)
        {
            if(str_contains($check_col,"-"))
            {
                $split = explode("-",$check_col);
                if(is_array($split))
                {
                    $valid = TRUE;
                    foreach($split as $key)
                    {
                        if(!ctype_alpha($key))
                        {
                            $valid = FALSE;
                        }

                    }
                    if ($valid)
                    {

                        $u_num = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($split[0]);
                        $l_num = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($split[1]);
                        if($u_num > $l_num)
                        {
                            for( $i = $l_num; $i <= $u_num; $i++)
                            {
                                $e_columns[]= $i;
                            }
                        }
                        else
                        {
                            for($i = $u_num; $i <= $l_num; $i++)
                            {
                                $e_columns[]= $i;
                            }
                        }
                    }
                }
            }
            else if (ctype_alpha($check_col))
            {
                $e_columns[] = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($check_col);
            }
            else
            {
                $col_null = TRUE;
            }
        }      
    }
    $dejavu = 0;
    foreach ($answers as $key => $value)
    {
        $j = 0;
            foreach ($value as $fields)
            {
                

                if(is_array($check_answ))
                {
                    $erase = FALSE;
                    
                    foreach($check_answ as $negate)
                    {
                        
                        if (!strcmp($fields,$negate) )
                        {
                            
                            $erase = TRUE; 
                            
                        }
                    }
                    if(!$erase || $answer_null)
                    {
                        if($fields)
                        {
                            $answers3[$key][$j] = $fields;  
                        }
                    }
                }
                else
                {
                    if(!(strcmp($fields,$negate) )|| $answer_null)
                    {
                        if($fields)
                        {
                            $answers3[$key][$j] = $fields;  
                        }
                    }
                }
                
                $z++;
                $j++;
                
            }


        $g++;
    }
    $over = FALSE;
    if(!is_null($answers3))
    {
        $answers4 = array();
        foreach ($answers3 as $key => $value)
        {
            $answers4[$key] = array_intersect($answers2[$key],$value);
        }
        $answers3 = $answers4;
        
        
    }

    $count = array();
    if(!is_null($answers3))
    {
        foreach ($answers4 as $key => $value)
        {
            foreach ($value as $key2 => $value2)
            {
                if (!empty($count[$key][$value2]))
                {
                    $count[$key][$value2]++;
                } 
                else    
                {
                    $count[$key][$value2] = 1;
                }
            }
        }
    }
    if(!is_null($answers3))
    {
        $answers4 = array();
        foreach ($answers3 as $key => $value)
        {
            $answers4[$key] = array_intersect($answers2[$key],$value);
        }
        $answers3 = $answers4;
    }

    $count = array();
    if(!is_null($answers3))
    {
        foreach ($answers4 as $key => $value)
        {
            foreach ($value as $key2 => $value2)
            {
                if (!empty($count[$key][$value2]))
                {
                    $count[$key][$value2]++;
                } 
                else    
                {
                    $count[$key][$value2] = 1;
                }
            }
        }
    }
    if(!is_null($answers3))
    {
        $answers3= multi_unique($answers3);
    }

    $went = FALSE;
    if(!is_null($answers3))
    {
        ob_end_clean();
                $fp = fopen('php://output','w');
        $fz = fopen('php://output', 'w');

        header('Content-Description: File Transfer');
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename="file.csv"');
        header('Content-Transfer-Encoding: binary');
        header('Expires: 0');
        header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
        header('Pragma: public');


        header('Content-Disposition: attachment; filename="filre.csv"');
        

        $i = 1;
        foreach ($answers3 as $key => $value)
        {
            $v_col = TRUE;
            foreach($e_columns as $obj)
            {
                if ($obj == $i)
                {
                    $v_col = FALSE;
                }
            }
            
            $list = array (array (mb_convert_encoding("{$key}", 'UTF-16LE', 'UTF-8')), array(mb_convert_encoding("Answers", 'UTF-16LE', 'UTF-8'),mb_convert_encoding("Frequency", 'UTF-16LE','UTF-8'),mb_convert_encoding("Percentage", 'UTF-16LE','UTF-8')));
            

            $total[$i] = 0;
            foreach ($value as $key2 => $value2)
            {
                $total[$i] += $count[$key][$value2];
            }
            foreach ($value as $key2 => $value2)
            {
                $percent = ($count[$key][$value2]/$total[$i]) * 100;
                $list2 = array (array (mb_convert_encoding("{$value2}",'UTF-16LE', 'UTF-8'),mb_convert_encoding("{$count[$key][$value2]}",'UTF-16LE', 'UTF-8'),mb_convert_encoding("{$percent}",'UTF-16LE', 'UTF-8')));
                foreach($check_quest as $item)
                {
                    if($v_col || $col_null)
                    {
                        if (trim($list2[0][0]) == trim(mb_convert_encoding("{$item}", 'UTF-16LE','UTF-8')) || ($quest_null))
                        {
                            $went = TRUE;
                            if ($dejavu != $list)
                            {
                                foreach ($list as $fields) 
                                {
                                    fputcsv($fp, $fields);
                                }
                                $dejavu = $list;
                            }


                            foreach ($list2 as $fields) 
                            {
                                fputcsv($fp, $fields);
                            }  
          

                        }
                    }
                }


               


     
            }
            $list3 = array (array ());
            

            foreach ($list3 as $fields) 
            {
                if($went)
                {
                    fputcsv($fp, $fields);
                }
            }  
            $went = FALSE;
            $i++;
        }
        fclose($fp);  
        fclose($fz);
        exit();
        return true;
    }
    else
    {
        $error = <<<DELIMETER
        <div class="error">"Every question and answer deleted. Could not generate csv file"</div>
        DELIMETER;
        echo $error;
        return false;
    }
}
    
    


//multi_ unique created by SML from https://stackoverflow.com/questions/37791484/remove-duplicate-values-from-multidimensional-3levels-array-with-php
function multi_unique($array){
        $temp = array_intersect_key($array, array_map("serialize", $array));
        foreach ($temp as $key => $value){
            if (is_array($value)){
               $temp[$key] = multi_unique($value);
            }
//the else if statement is to counter the problem where the deepest level of each path is different
            else if (count($temp) == count($temp, COUNT_RECURSIVE)){
                $temp=array_unique($temp);
            }
        }
        return $temp;
    }
?>