<?php
/**
 * Created by PhpStorm.
 * User: Administrator
 * Date: 2018/3/3 0003
 * Time: 11:14
 */

namespace Home\Controller;
use Home\Model\PHPExcelModel;
use Think\Controller;

class TestController extends Controller
{
    /**
     * @throws \PHPExcel_Exception
     */
    public function index(){
        $excel =new PHPExcelModel();
        $excel->setConf(array(
            'SheetList'=>array(
                'A', 'B', 'C', 'D',),
            'fileName'=>"测试".time(),
            'path'=>"./Public/excel/",
        ));
        $arr = array(
            array("format"=>'int',"width"=>30,"liteName"=>"ID"),
            array("format"=>'text',"width"=>30,"liteName"=>"number"),
            array("format"=>'text',"width"=>30,"liteName"=>"name"),
            array("format"=>'text',"width"=>30,"liteName"=>"order"),

        );
        $data = array(
            array('1','564646','asdad','54464644'),
            array('2','564646','asdad','54464644'),
            array('3','564646','asdad','54464644'),
            array('4','564646','asdad','54464644'),
            array('5','564646','asdad','54464644'),
            array('6','564646','asdad','54464644'),
            array('9','564646','asdad','54464644'),
        );
        $excel->setActiveSheet($arr,"测试标题");
        $excel->foundExcelFile();
        $excel->excelWRDate($data);
    }
}


