<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2018/6/4 0004
 * Time: 10:05
 */

namespace Home\Model;

import("Vendor.PHPExcel.PHPExcel");
import("Vendor.PHPExcel.PHPExcel.Reader.Excel2007");
import("Vendor.PHPExcel.PHPExcel.IOFactory");
use PHPExcel;
use Think\Model;

class MyExcelModel extends Model
{
    Protected $autoCheckFields = false;
    private $objPHPExcel;
    private $objWriter;
    private $objActSheet;
    private $writerType='Excel2007';//Excel版本（'Excel2007'）
    private $sheets;
    protected  $List=array();//初始化列
    protected  $StarRow=1;//开始行
    protected $filePathName;//生成文件的文件名和地址
    protected $fileName;//创建文件的文件名
    protected $path='./public';//创建文件的路径
    protected $ActiveSheetIndex;//
    private $rowKeys;//标识指向空白行


    /**
     * MyExcelModel constructor.
     * @param $conf
     * @throws \PHPExcel_Exception
     */
    public function __construct($conf)
    {
        parent::__construct();
        //设置工作表数量
        $this->sheets=($conf['sheets']&&is_array($conf['sheets']))? $conf['sheets']:
            array(
                array(
                    'title'=>array(
                        'name'=>'Sheet 1',
                        'fondSize'=>'18',
                        ),
                    //以下只供测试
                    'field'=>array(
                            array(
                                'name'=>'测试',
                                'textType'=>'text',
                                'width'=>'50',
                            ),
                            array(
                                'name'=>'测试',
                                'textType'=>'text',
                                'width'=>'50',
                            ),
                        )
                    ),
                );

        //
        //新建一个Excel对象
        $this->objPHPExcel = new PHPExcel();
        //设置Excel的Sheet数量和标题
        $this->setSheets();
        //选择设置的Excel的Sheet对象
    }


    /**
     * 设置Excel的Sheet数量和标题
     * @throws \PHPExcel_Exception
     */
    private function setSheets(){
        //用于sheet计数
        $num=0;
        foreach ($this->sheets as $k=>$v){

            //第一个sheet不需要创建，之后每添加一个都需要创建新的工作空间
            if($num>0)$this->objPHPExcel->createSheet($k);

            //设置sheet的标识，（设置为对应数组的key）
            $this->objPHPExcel->setActiveSheetIndex($k);

            //获取当前工作表
            $sheet = $this->objPHPExcel->getActiveSheet();
            if($v['title']){

                //设置sheet的标题
                $sheet->setTitle($v['title']['name']);

                //合并当前第一行
                $sheet->mergeCells(
                    "A1:{$this->List[count($this->List)-1]}1"
                );

                //设置字体样式
                $sheet->getStyle('A1')
                    ->getFont()
                    ->setBold(true);
                $sheet->getStyle('A1')
                    ->getFont()
                    ->setSize($v['title']['fondSize']?$v['title']['fondSize']:13);

                //把标识指向空白行
                $this->rowKeys = array("$k"=>2);
            }
            $this->rowKeys[$k]?'':$this->rowKeys[$k]=1;

            //设置所有文本居中
            $sheet->getStyle()
                ->getAlignment()
                ->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_RIGHT
                );

            //判断是否存在表格字段设置
            if($v['field']){
                foreach ($v['field'] as $fk => $fv){

                    //设置字体大小
                    $sheet->getStyle('A1')
                        ->getFont()
                        ->setSize($fv['textType']?$fv['textType']:12);

                    //内容自适应
                    $sheet->getColumnDimension( $this->List[$fk].$this->rowKeys[$k])
                        ->setAutoSize(true);

                    //设置字段宽度
                    $sheet->getColumnDimension( $this->List[$fk].$this->rowKeys[$k])
                        ->setWidth($fv['width']);

                    //设置单元格文本格式
                    $sheet->getStyle($this->List[$fk].$this->rowKeys[$k])->getNumberFormat()
                        ->setFormatCode($this->My_Get_Style($fv['textType']));
                }

                //把标识指向空白行
                $this->rowKeys[$k]++;

                //如果存在标题就给标题赋值
                if($v['title']){
                    $sheet->setCellValue('A1',$v['title']['name']);
                }

                //设置字段值
                foreach ($v['field'] as $fk => $fv){
                    $sheet->setCellValue( $this->List[$fk].$this->rowKeys[$k],$fv['name']);//插入单元格数据
                }
            }
            $num++;
        }

    }

    public function excelWRDate($data,$k){
        $objReader =PHPExcel_IOFactory:: createReader('Excel2007' );
        $objPHPExcel =$objReader->load( "$this->filePathName" );
        $this->objPHPExcel->setActiveSheetIndex($k);
//吧数组的内容从A2开始填充
//        $dataArray= array( array("2010" ,    "Q1",  "UnitedStates",  790),
//            array("2010" ,    "Q2",  "UnitedStates",  730),
//        );


        $objPHPExcel->getActiveSheet()->fromArray($data, NULL, 'A'.$this->rowKeys[$k]);
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

    /**
     * 获取单元格格式
     * @param $type
     * @return string
     */
    private function My_Get_Style($type){
        switch ($type){
            case 'text':return \PHPExcel_Style_NumberFormat::FORMAT_TEXT;
            case 'int':return \PHPExcel_Style_NumberFormat::FORMAT_NUMBER;
            case 'float':return \PHPExcel_Style_NumberFormat::FORMAT_NUMBER_00;
            case 'price':return \PHPExcel_Style_NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1;
            default:return \PHPExcel_Style_NumberFormat::FORMAT_TEXT;
        }
    }
}
