<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2018/3/21 0021
 * Time: 13:45
 */
namespace Home\Model;
use PHPExcel_IOFactory;
import("Vendor.PHPExcel.PHPExcel");
import("Vendor.PHPExcel.PHPExcel.Reader.Excel2007");
import("Vendor.PHPExcel.PHPExcel.IOFactory");
//require_once '../Classes/PHPExcel/IOFactory.php';
//require_once '../Classes/PHPExcel/IOFactory.php';
class PHPExcelModel extends \Think\Model{
    Protected $autoCheckFields = false;
    private $objPHPExcel;
    private $objWriter;
    private $objActSheet;
    private $writerType='Excel2007';//Excel版本（'Excel2007'）
    protected  $SheetList=array();//初始化列
    protected  $SheetStarRow=1;//开始行
    protected $filePathName;//生成文件的文件名和地址
    protected $fileName;//创建文件的文件名
    protected $path='./public';//创建文件的路径

    /**
     * PHPExcelModel constructor.
     * @throws \PHPExcel_Exception
     */
    public function __construct()
    {
        parent::__Construct();
        $this->objPHPExcel = new \PHPExcel();
        $this->objActSheet = $this->objPHPExcel->getActiveSheet();//单元格属性
        $this->objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, $this->writerType);
    }
    public function setConf($conf){
        $arr = array(
            'A', 'B', 'C', 'D', 'E', 'F', 'G',
            'H', 'I', 'J', 'K', 'L', 'M',
            'N', 'O', 'P', 'Q', 'R', 'S',
            'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
        );
        $this->SheetList = $conf['SheetList']?$conf['SheetList']:$arr;
        $this->fileName = $conf['fileName']? $conf['fileName']:time()."未定义文件名";
        $conf['SheetStarRow']?$this->SheetStarRow = $conf['SheetStarRow']:'';
        $conf['path']?$this->path = $conf['path']:'';


    }
    /**
     * 设置标题设置列表样式
     * @param $arr
     * @param $sheetName
     * @throws \PHPExcel_Exception
     */
    public function setActiveSheet($arr, $sheetName){
        //设置标题
        $this->objActSheet
            ->mergeCells(
                "{$this->SheetList[0]}{$this->SheetStarRow}:{$this->SheetList[count($this->SheetList)-1]}{$this->SheetStarRow}"
            );//'A1:B1'
        $this->objActSheet->setCellValue('A'.$this->SheetStarRow,$sheetName);
        $this->objActSheet->getStyle('A'.$this->SheetStarRow)->getFont()->setBold(true);
        $this->objActSheet->getStyle('A'.$this->SheetStarRow)->getFont()->setSize(18);
        //设置列表样式
        foreach ($arr as $a=>$data){
            $this->objActSheet->getColumnDimension( $this->SheetList[$a])->setAutoSize(true);   //内容自适应
            $this->objActSheet->getColumnDimension( $this->SheetList[$a])->setWidth($a['width']);        //30宽
            $this->objActSheet->getStyle($this->SheetList[$a])->getNumberFormat()
                ->setFormatCode($this->getSheet($a['format']));
            $this->objActSheet->setCellValue($this->SheetList[$a].$this->SheetStarRow,$arr['liteName']);
        }
    }

    /**
     * 获取单元格格式
     * @param $type
     * @return string
     */
    private function getSheet($type){
        switch ($type){
            case 'text':return \PHPExcel_Style_NumberFormat::FORMAT_TEXT;
            case 'int':return \PHPExcel_Style_NumberFormat::FORMAT_NUMBER;
            case 'float':return \PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00;
            case 'price':return \PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1;
            default:return \PHPExcel_Style_NumberFormat::FORMAT_TEXT;
        }
    }

    /**
     * 创建文件（excel）
     * @param string $path 路径
     */
    public function foundExcelFile(){
        if (is_null($this->fileName)) {
            $savefile = time();
        } else {
            //防止中文命名，下载时ie9及其他情况下的文件名称乱码
            iconv('UTF-8', 'GB2312', $this->fileName);
        }
        try {

            $filePathName = $this->path.$this->fileName.time().'.xlsx';
            $this->objWriter->save($filePathName);
        } catch (\PHPExcel_Reader_Exception $e) {
            error('SERVER_BUSY',$e);
        } catch (\PHPExcel_Writer_Exception $e) {
            error('SERVER_BUSY',$e);
        }
        $this->filePathName=$filePathName;
    }

    /**
     * 写入数据到现有表中
     * @param $data
     * @param int $baseRow
     * @throws \PHPExcel_Exception
     */
    public function excelWRDate($data, $baseRow=3){
        try {
            //从文件中加载PHPExcel，使用自动phpexcelreaderireader解析
            $objPHPExcel = \PHPExcel_IOFactory::load($this->filePathName);
        } catch (\PHPExcel_Reader_Exception $e) {
            error('SERVER_BUSY',$e);
        }
        foreach($data as $r => $dataRow){
            $row= $baseRow +$r;//$row是循环操作行的行号
            $objPHPExcel->getActiveSheet()->insertNewRowBefore($row,1);//插入新行
            //遍历一条数据中的每一个列
            foreach ($dataRow as $b){
                $objPHPExcel->getActiveSheet()->setCellValue( $this->SheetList[$r].$row,$b);//插入单元格数据
            }
        }
    }

}