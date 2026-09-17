 call doex("create database if not exists `testDB` character set latin1 collate latin1_german2_ci;",0)
 call doex("grant all privileges on `testDB`.* to 'praxis'@'%' identified by 'sonne' with grant option",0)
 call doex("grant all privileges on `testDB`.* to 'praxis'@'localhost' identified by 'sonne' with grant option",0)
 call doex("use `testDB`",0)
