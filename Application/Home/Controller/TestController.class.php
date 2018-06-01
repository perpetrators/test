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
                'A', 'B', 'C', 'D', 'E', 'F', 'G',
                'H', 'I', 'J', 'K', 'L', 'M',
                'N', 'O', 'P', 'Q', 'R', 'S',),
            'fileName'=>"测试".time(),
            'path'=>"",
        ));
        $arr = array(
            array("format"=>'int',"width"=>30,"liteName"=>"ID"),
            array("format"=>'text',"width"=>30,"liteName"=>"number"),
            array("format"=>'text',"width"=>30,"liteName"=>"name"),
            array("format"=>'text',"width"=>30,"liteName"=>"order"),

        );
        $excel->setActiveSheet($arr,"测试标题");
        $excel->foundExcelFile();
    }
}


