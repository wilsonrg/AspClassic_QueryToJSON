CREATE TABLE `login` 
( `id` int(11) NOT NULL AUTO_INCREMENT,
  `login` varchar(15) NOT NULL,
  `senha` varchar(15) NOT NULL,
   PRIMARY KEY (`id`) 
) ENGINE=InnoDB AUTO_INCREMENT=9 DEFAULT CHARSET=utf8;

-- Inserindo informações na tabela login para serem recuperadas depois
INSERT INTO login(login,senha) VALUES('login01','senha01'),('login02','senha02');