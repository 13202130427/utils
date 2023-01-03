<?php

namespace Uroad\Utils\Mysql;

class Query
{
    public static $instance = null;
    private static $sql;
    private static $table;
    private static $tableRename;
    private static $field;
    private static $join;
    private static $where;
    private static $groupBy;
    private static $orderBy;


    public function __construct()
    {
        self::$sql = "";
        self::$table = [];
        self::$field = [];
        self::$join = [];
        self::$where = ['TRUE'];
        self::$groupBy = "";
        self::$orderBy = "";
    }

    /**
     * @param string|array $table 表名
     */
    public function table($table)
    {
        if (is_string($table)) {
            $table = explode(',', $table);
        }
        self::$table = array_merge(self::$table,$table);
        return $this;
    }

    public function select($field)
    {
        if (is_string($field)) {
            $field = explode(',', $field);
        }
        self::$field = $field;
        return $this;
    }

    /**
     * @param string|array $cond JOIN连接语句
     */
    public function join($cond)
    {
        if (is_string($cond)) {
            $cond = explode(',', $cond);
        }
        foreach ($cond as $value) {
            self::$join[] = $value;
        }
        return $this;
    }

    public function where($sql,$param = [],$type = 'AND')
    {
        if (!in_array($type,['AND','OR'])) {
            throw new \Exception('不符合的关联关系');
        }
        if (mb_substr_count($sql,'?') !== count($param)) {
            throw new \Exception('参数有误!');
        }
        foreach ($param as $value) {
            $index = stripos($sql,'?');
            if ($index !== false) {
                if (is_string($value)) $value = " '".$value."' ";
                $sql = substr_replace($sql,$value,$index,1);
            }
        }
        self::$where[] = $type .' '.$sql;
        return $this;
    }

    public function groupBy($cond)
    {
        if (is_string($cond)) {
            $cond = explode(',', $cond);
        }
        foreach ($cond as $value) {
            self::$groupBy[] = $value;
        }
        return $this;
    }

    public function orderBy($cond)
    {
        if (is_string($cond)) {
            $cond = explode(',', $cond);
        }
        foreach ($cond as $value) {
            self::$orderBy[] = $value;
        }
        return $this;
    }

    public function as($name)
    {
        self::$tableRename = $name;
        return $this;
    }


    public function get()
    {
        try {
            if (empty(self::$table)) throw new \Exception('未指定表');
            if (empty(self::$field)) throw new \Exception('未指定查询字段');
            if (count(self::$table) == 1) {
                //单表查询
                $this->queryToSingle();
            } else {
                //分表查询
                $this->queryToMore();
            }
            if (self::$join) self::$sql .= "\n" . implode("\n", self::$join);
            self::$sql .= ' WHERE '. "\n" . implode("\n", self::$where);
            if (self::$groupBy) self::$sql .= 'GROUP BY ' .implode(',',self::$groupBy);
            if (self::$orderBy) self::$sql .= 'ORDER BY ' .implode(',',self::$orderBy);
            return self::$sql;
        }catch (\Throwable $exception) {
            throw new \Exception($exception);
        }
    }

    protected function queryToSingle()
    {
        self::$sql = 'SELECT '.implode(',',self::$field) . ' FROM '.self::$table;
        if (self::$tableRename) self::$sql .= ' AS '.self::$tableRename;
    }

    protected function queryToMore()
    {
        if (empty(self::$tableRename)) throw new \Exception('多表联合查询必须指定主表别名');
        self::$sql = 'SELECT '.implode(',',self::$field) . ' FROM (';

        if (empty(self::$table)) throw new \Exception('未指定表');
        $contentSql = [];
        foreach (self::$table as $table) {
            $sql = 'SELECT '.self::$tableRename.'.* FROM '.$table .' AS '.self::$tableRename;
            if (self::$join) $sql .= "\n" . implode("\n", self::$join);
            $sql .= ' WHERE '. "\n" . implode("\n", self::$where);
            $contentSql[] = $sql;
        }
        self::$sql .= implode(' UNION ',$contentSql);
        self::$sql .= ') '.self::$tableRename;
    }
}
