<?php

use Uroad\Utils\File\Excel;

class Test {

    public function write()
    {
        $excel = Excel::load(1);
        $fileUrl = './a.xlsx';
        $header = [
            [
                'A' => ['title'=>'防御工作及受灾情况调查统计表','value' => '','merge_column'=>'N']
            ],
            [
                'E' => ['title'=>'防御工作','value' => '','merge_column'=>'H'],
                'I' => ['title'=>'受灾情况','value' => '','merge_column'=>'L']
            ],
            [
                'A' => ['title'=>'序号','value' => 'A1','merge_line'=>'1'],
                'B' => ['title'=>'二级单位','value' => 'B1','merge_line'=>'1'],
                'C' => ['title'=>'基层单位','value' => 'C1','merge_line'=>'1'],
                'D' => ['title'=>'联络人','value' => 'D1','merge_line'=>'1'],
                'E' => ['title'=>'值守人员','value' => 'E1'],
                'F' => ['title'=>'车辆设备','value' => 'F1'],
                'G' => ['title'=>'转移设备','value' => 'G1'],
                'H' => ['title'=>'其他重点工作','value' => 'H1'],
                'I' => ['title'=>'人员伤亡情况','value' => 'I1'],
                'J' => ['title'=>'灾害简要描述','value' => 'J1'],
                'K' => ['title'=>'经济损失','value' => 'K1'],
                'L' => ['title'=>'主要损失情况','value' => 'L1'],
                'M' => ['title'=>'下一步重点工作计划举措','value' => 'M1',''],
                'N' => ['title'=>'其他','value' => 'N1'],
            ]
        ];
        $data = [
          [
              'A1' => 'a',
              'B1' => 'b',
              'C1' => 'c',
              'D1' => 'd',
              'E1' => 'e',
              'F1' => 'f',
              'G1' => 'g',
              'H1' => 'h',
              'I1' => 'i',
              'J1' => 'j',
              'K1' => 'k',
              'L1' => 'l',
              'M1' => 'm',
              'N1' => 'n',
          ]
        ];
        try {
            $url = $excel->setHeader($header)->setData($data)->generate($fileUrl);
            var_dump('生成成功');
            var_dump($url);die;
        } catch (\Exception $e) {
            var_dump($e);die;
        }
    }


    public function read()
    {
        $excel = Excel::load();
        $fileUrl = './a.xlsx';
        $header = [
            [
                'A' => ['title'=>'防御工作及受灾情况调查统计表','value' => '','merge_column'=>'N']
            ],
            [
                'E' => ['title'=>'防御工作','value' => '','merge_column'=>'H'],
                'I' => ['title'=>'受灾情况','value' => '','merge_column'=>'L']
            ],
            [
                'A' => ['title'=>'序号','value' => 'A1','merge_line'=>'1'],
                'B' => ['title'=>'二级单位','value' => 'B1','merge_line'=>'1'],
                'C' => ['title'=>'基层单位','value' => 'C1','merge_line'=>'1'],
                'D' => ['title'=>'联络人','value' => 'D1','merge_line'=>'1'],
                'E' => ['title'=>'值守人员','value' => 'E1'],
                'F' => ['title'=>'车辆设备','value' => 'F1'],
                'G' => ['title'=>'转移设备','value' => 'G1'],
                'H' => ['title'=>'其他重点工作','value' => 'H1'],
                'I' => ['title'=>'人员伤亡情况','value' => 'I1'],
                'J' => ['title'=>'灾害简要描述','value' => 'J1'],
                'K' => ['title'=>'经济损失','value' => 'K1'],
                'L' => ['title'=>'主要损失情况','value' => 'L1'],
                'M' => ['title'=>'下一步重点工作计划举措','value' => 'M1'],
                'N' => ['title'=>'其他','value' => 'N1'],
            ]
        ];
        try {
            $data = $excel->dataConvert($header, $fileUrl,'Xlsx');
            var_dump('读取数据');
            var_dump($data);die;
        } catch (\Exception $e) {
            var_dump($e);die;
        }
    }
}