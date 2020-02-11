<?php
class DatabaseClass {
	var $hostname = "localhost"; 	 // database hostname
	var $username = "root";			 //database user name
	var $password = ""; 			 // database pass
	var $database = "recepdur_vtr";	 //database name 

function connect(){
	$this->conn = mysql_connect($this->hostname, $this->username, $this->password) or header("Location: 404.html");
	mysql_query("SET NAMES 'utf8'");
	mysql_query("SET CHARACTER SET 'utf8'");
	mysql_query("SET COLLATION_CONNECTION = 'utf8_turkish_ci'");
	$select_db = mysql_select_db($this->database) or header("Location: 404.html");
}

	function query($a){
		return mysql_query($a,mysql_connect($this->hostname, $this->username, $this->password)) ;
	}
    function farray($result)
	{
		return mysql_fetch_array($result);
	}
	
	function fassoc($result){
		return mysql_fetch_assoc($result);
	}
		
	function numrows($result){
		return mysql_num_rows($result);
	}
	
	function affected(){
		return mysql_affected_rows();
	}
 
    function free($result){
		return mysql_free_result($result);
	}
	
	function insert(){
		return mysql_insert_id();
	}
	
	function result($result){
		return mysql_result($result,0);
	}
	
	function close(){
		return mysql_close($this->conn);
	}
}
	$DB = new DatabaseClass;
	$DB->connect();
?>