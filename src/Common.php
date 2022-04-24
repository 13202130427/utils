<?php

namespace Uroad\Utils;

class Common
{
    public static function isEmpty(&$param,$field)
    {
        if (!isset($param[$field]) || $param[$field] == '' || $param[$field] == [] || $param[$field] == null) {
            return true;
        }
        return false;
    }

    /**
     * 若查找数组指定键值为空且写入数组指定键值不存在 写入数组 指定键值 插入 默认值
     * 若查找数组指定键值为空且写入数组指定键值存在 不操作
     * 若查找数组指定键值存在 写入数组 指定键值 插入 查找数组数据
     * @param $findParam array 查找数组
     * @param $findField string 查找键值
     * @param $writeParam array 写入数组
     * @param $writeField string 写入键值
     * @param $value string 默认值
     */
    public static function isEmptySetValue(&$findParam,$findField,&$writeParam,$writeField,$value)
    {
        if (!isset($findParam[$findField]) || $findParam[$findField] == '' || $findParam[$findField] == [] || $findParam[$findField] == null) {
            if (!isset($writeParam[$writeField]) || $writeParam[$writeField] == '' || $writeParam[$writeField] == [] || $writeParam[$writeField] == null) {
                $writeParam[$writeField] = $value;
                return;
            }
            return;
        }
        $writeParam[$writeField] = $findParam[$findField];
    }
}