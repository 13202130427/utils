<?php

namespace Uroad\Utils\Mysql;

/**
 * @method static table(string[] $array)
 */
class Db
{
    public static function __callStatic($name, $arguments)
    {
        return call_user_func_array([new Query(),$name],[$arguments[0]]);
    }

}
