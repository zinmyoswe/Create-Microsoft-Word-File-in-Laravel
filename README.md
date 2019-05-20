# Create-Microsoft-Word-File-in-Laravel
Create and Print Microsoft Word File, ODF, HTML File in Laravel


If you are saving a document as an `Word File`. Edit in `DocumentController.php`

```php
    public function store(Request $request)
    {
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $phpWord->addSection();
        $text = $section->addText($request->get('name'));
        $text = $section->addText($request->get('email'));
        $text = $section->addText($request->get('number'),array('name'=>'Arial','size' => 20,'bold' => true));
        $section->addImage("./images/python.png",array('width'=>'300','padding' => '20px'));  
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save('zinmyosweSE.docx');
        return response()->download(public_path('zinmyosweSE.docx'));
    }
```
If you are saving a document as an `ODF File`. Edit in `DocumentController.php`
```php
public function store(Request $request)
    {
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $phpWord->addSection();
        $text = $section->addText($request->get('name'));
        $text = $section->addText($request->get('email'));
        $text = $section->addText($request->get('number'),array('name'=>'Arial','size' => 20,'bold' => true));
        $section->addImage("./images/Krunal.jpg");  
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'ODText');
        $objWriter->save('Appdividend.odt');
        return response()->download(public_path('Appdividend.odt'));
    }
```
If you are saving a document as an `HTML File`. Edit in `DocumentController.php`
```php
public function store(Request $request)
    {
        $phpWord = new \PhpOffice\PhpWord\PhpWord();
        $section = $phpWord->addSection();
        $text = $section->addText($request->get('name'));
        $text = $section->addText($request->get('email'));
        $text = $section->addText($request->get('number'),array('name'=>'Arial','size' => 20,'bold' => true));
        $section->addImage("./images/Krunal.jpg");  
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML');
        $objWriter->save('Appdividend.html');
        return response()->download(public_path('Appdividend.html'));
    }
```
