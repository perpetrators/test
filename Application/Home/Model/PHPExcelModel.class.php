<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2018/3/21 0021
 * Time: 13:45
 */
namespace Home\Model;
use PHPExcel_IOFactory;
use Think\Model;

import("Vendor.PHPExcel.PHPExcel");
import("Vendor.PHPExcel.PHPExcel.Reader.Excel2007");
import("Vendor.PHPExcel.PHPExcel.IOFactory");
//require_once '../Classes/PHPExcel/IOFactory.php';
//require_once '../Classes/PHPExcel/IOFactory.php';
class PHPExcelModel extends Model{
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
        $this->objPHPExcel->setactivesheetindex(0);
        $this->objPHPExcel->getActiveSheet()->setTitle('标题1');
        $this->objActSheet = $this->objPHPExcel->getActiveSheet();//单元格属性




        $this->objPHPExcel->createSheet();
        $this->objPHPExcel->setActiveSheetIndex(1);
        $this->objPHPExcel->getActiveSheet()->setTitle('标题2');


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
//        $conf['SheetStarRow']?$this->SheetStarRow = $conf['SheetStarRow']:'';
        $conf['path']?$this->path = $conf['path']:'';


    }
    /**
     * 设置标题设置列表样式
     * @param $arr
     * @param $sheetName
     * @throws \PHPExcel_Exception
     */
    public function setActiveSheet($arr, $sheetName){
//设置当前的sheet\


        $this->objActSheet->getStyle()->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);    //水平方向上对齐
        //设置标题
        $this->objActSheet
            ->mergeCells(
                "{$this->SheetList[0]}{$this->SheetStarRow}:{$this->SheetList[count($this->SheetList)-1]}{$this->SheetStarRow}"
            );//'A1:B1'

        $this->objActSheet->getStyle('A'.$this->SheetStarRow)->getFont()->setBold(true);
        $this->objActSheet->getStyle('A'.$this->SheetStarRow)->getFont()->setSize(18);
        $this->objActSheet->setCellValue('A'.$this->SheetStarRow,$sheetName);
        $this->objActSheet->setCellValueExplicit("A$this->SheetStarRow", $sheetName,\PHPExcel_Cell_DataType::TYPE_STRING);
        //设置列表样式
        foreach ($arr as $a=>$data){
            $this->objActSheet->getColumnDimension( $this->SheetList[$a])->setAutoSize(true);   //内容自适应
            $this->objActSheet->getColumnDimension( $this->SheetList[$a])->setWidth($data['width']);        //30宽
            $this->objActSheet->getStyle($this->SheetList[$a])->getNumberFormat()
                ->setFormatCode($this->getSheet($a['format']));
            $this->objActSheet->setCellValue($this->SheetList[$a].$this->SheetStarRow,$arr['liteName']);
//            $this->objActSheet->setCellValueExplicit("$letter[$i]2", $value['title'],\PHPExcel_Cell_DataType::TYPE_STRING);

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
     * @throws \PHPExcel_Reader_Exception
     */
    public function foundExcelFile(){
        if (is_null($this->fileName)) {
            $savefile = time();
        } else {
            //防止中文命名，下载时ie9及其他情况下的文件名称乱码
            iconv('UTF-8', 'GB2312', $this->fileName);
        }
        $this->objPHPExcel->setActiveSheetIndex(0)->setCellValue('A'.$this->SheetStarRow,'sdadadasd');
//        $this->objActSheet->setCellValueExplicit("A$this->SheetStarRow", 'asdadas',\PHPExcel_Cell_DataType::TYPE_STRING);

        try {

            $this->objWriter = \PHPExcel_IOFactory::createWriter($this->objPHPExcel, 'Excel2007');
            $filePathName = $this->path.$this->fileName.time().'.xls';
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
    public function excelWRDate($data, $baseRow=2){
        $objReader =PHPExcel_IOFactory:: createReader('Excel2007' );
        $objPHPExcel =$objReader->load( "$this->filePathName" );
        $this->objPHPExcel->setActiveSheetIndex(0);
//吧数组的内容从A2开始填充
//        $dataArray= array( array("2010" ,    "Q1",  "UnitedStates",  790),
//            array("2010" ,    "Q2",  "UnitedStates",  730),
//        );


        $objPHPExcel->getActiveSheet()->fromArray($data, NULL, 'A2');
//        foreach($data as $r => $dataRow){
//            $row= $baseRow +$r;//$row是循环操作行的行号
//            $sheet = $objPHPExcel->getActiveSheet();
//            $sheet->insertNewRowBefore($row,1);
//            //遍历一条数据中的每一个列
//            foreach ($dataRow as $k=>$b){
//                $sheet->setCellValue( $this->SheetList[$k].$row,$b);//插入单元格数据
//            }
//        }
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $this->writerType);
        $objWriter->save($this->filePathName);
    }

}