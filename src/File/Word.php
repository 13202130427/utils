<?php


namespace Uroad\Utils\File;


use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\PhpWord;

class Word
{
    public static $instance = null;
    private static $obj;

    private function __construct($config = [])
    {
        if (!empty($config)) {

        }
        self::$obj = new PhpWord();
    }

    private function __clone()
    {
        // TODO: Implement __clone() method.
    }


    /**
     * 加载入口
     * @param array $config
     * @return Word
     */
    public static function load($config = [])
    {
        if (!(self::$instance instanceof self)) {
            self::$instance = new self($config);
        }
        return self::$instance;
    }

    public function addPage()
    {
        return self::$obj->addSection();
    }

    /**
     * @param  $section
     * @param $title
     * @param $data
     * @param int[] $config
     */
    public function addTable($section,$title,$data,$config = ['size'=>16])
    {
        $section->addText('Basic table',$config);
        $table = $section->addTable();
        //设置表头
        $titleRow = [];
        $table->addRow();
        foreach ($title as $key=>$value) {
            $table->addCell(1750)->addText($value);
            array_push($titleRow,$key);
        }
        //设置表数据
        foreach ($data as $rowData) {
            $table->addRow();
            foreach ($titleRow as $item) {
               $table->addCell(1750)->addText($rowData[$item]);
            }
        }
    }

    /**
     * 生成表格
     * @param string $path 路程 存在即保存指定路径
     * @param string $format
     */
    public function generate($path ='',$format='Word2007')
    {
        $objWriter = IOFactory::createWriter(self::$obj,$format);
        if (!empty($path)) {
            $objWriter->save($path);
            return $path;
        }
        $file = 'test.docx';
        header("Content-Description: File Transfer");
        header('Content-Disposition: attachment; filename="' . $file . '"');
        header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        header('Content-Transfer-Encoding: binary');
        header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
        header('Expires: 0');
        $objWriter->save('php://output');
        exit;
    }

}