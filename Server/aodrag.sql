
-- Volcando estructura de base de datos para aodrag
CREATE DATABASE IF NOT EXISTS `aodrag` /*!40100 DEFAULT CHARACTER SET utf8 */;
USE `aodrag`;

-- Volcando estructura para tabla aodrag.castillo
CREATE TABLE IF NOT EXISTS `castillo` (
  `id` smallint(6) NOT NULL AUTO_INCREMENT,
  `nombre` varchar(25) DEFAULT NULL,
  `dueño` int(11) DEFAULT NULL,
  `fecha_conquista` datetime DEFAULT NULL,
  `mapa` smallint(5) unsigned DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.castillo: ~4 rows (aproximadamente)
/*!40000 ALTER TABLE `castillo` DISABLE KEYS */;
INSERT INTO `castillo` (`id`, `nombre`, `dueño`, `fecha_conquista`, `mapa`) VALUES
	(1, 'Castillo Norte', 0, '2018-12-25 12:55:21', 1),
	(2, 'Castillo Sur', 0, '2018-12-25 12:55:20', 1),
	(3, 'Castillo Este', 0, '2018-12-25 12:55:18', 1),
	(4, 'Castillo Oeste', 0, '2018-12-25 12:55:17', 1),
	(5, 'Fortaleza', 0, '2018-12-25 12:55:14', 1);
/*!40000 ALTER TABLE `castillo` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.clan
CREATE TABLE IF NOT EXISTS `clan` (
  `id` smallint(5) unsigned NOT NULL AUTO_INCREMENT,
  `fundador` smallint(5) unsigned DEFAULT NULL,
  `nombre` varchar(25) DEFAULT NULL,
  `fecha_fundacion` datetime DEFAULT NULL,
  `lider` mediumint(8) unsigned DEFAULT NULL,
  `desc` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.clan: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `clan` DISABLE KEYS */;
/*!40000 ALTER TABLE `clan` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.config_intervalo
CREATE TABLE IF NOT EXISTS `config_intervalo` (
  `User_AtacarMelee` smallint(6) DEFAULT NULL,
  `User_LanzarMagia` smallint(6) DEFAULT NULL,
  `NPC_AtacarMelee` smallint(6) DEFAULT NULL,
  `NPC_LanzarMagia` smallint(6) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.config_intervalo: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `config_intervalo` DISABLE KEYS */;
/*!40000 ALTER TABLE `config_intervalo` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.config_party
CREATE TABLE IF NOT EXISTS `config_party` (
  `party_porcentaje_exp_por_jugador` float(10,0) DEFAULT NULL,
  `party_distancia_maxima_para_exp` tinyint(10) DEFAULT NULL,
  `heartbeat` datetime DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- Volcando datos para la tabla aodrag.config_party: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `config_party` DISABLE KEYS */;
/*!40000 ALTER TABLE `config_party` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.cuenta
CREATE TABLE IF NOT EXISTS `cuenta` (
  `id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `password` varchar(32) DEFAULT NULL,
  `mail` varchar(50) DEFAULT NULL,
  `fecha_creacion` datetime DEFAULT current_timestamp(),
  `pregunta` varchar(100) DEFAULT NULL,
  `respuesta` varchar(100) DEFAULT NULL,
  `bloqueada` tinyint(1) DEFAULT 0,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.cuenta: ~2 rows (aproximadamente)
/*!40000 ALTER TABLE `cuenta` DISABLE KEYS */;
INSERT INTO `cuenta` (`id`, `password`, `mail`, `fecha_creacion`, `pregunta`, `respuesta`, `bloqueada`) VALUES
	(1, '1', '1', '2018-11-09 04:03:10', NULL, NULL, 0),
	(2, '2', '2', '2018-11-09 04:03:15', NULL, NULL, 0);
/*!40000 ALTER TABLE `cuenta` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.efecto
CREATE TABLE IF NOT EXISTS `efecto` (
  `id` smallint(6) unsigned NOT NULL AUTO_INCREMENT,
  `nombre` char(50) NOT NULL DEFAULT '0',
  `tipo` smallint(6) unsigned NOT NULL DEFAULT 0,
  `descripcion` varchar(50) NOT NULL DEFAULT '0',
  `duracion` mediumint(8) unsigned NOT NULL DEFAULT 0,
  `intervalo` smallint(6) unsigned NOT NULL DEFAULT 0,
  `limite` tinyint(4) unsigned NOT NULL DEFAULT 0,
  `limite_origen` tinyint(4) unsigned NOT NULL DEFAULT 0,
  `beneficioso` tinyint(1) unsigned NOT NULL DEFAULT 0,
  `grh` smallint(5) unsigned NOT NULL DEFAULT 0,
  `aplicado` tinyint(1) unsigned NOT NULL DEFAULT 0,
  `enviar_a_cliente` tinyint(1) unsigned NOT NULL DEFAULT 0,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.efecto: ~2 rows (aproximadamente)
/*!40000 ALTER TABLE `efecto` DISABLE KEYS */;
INSERT INTO `efecto` (`id`, `nombre`, `tipo`, `descripcion`, `duracion`, `intervalo`, `limite`, `limite_origen`, `beneficioso`, `grh`, `aplicado`, `enviar_a_cliente`) VALUES
	(1, 'Daño mágico', 0, 'Daño mágico ', 0, 0, 0, 0, 0, 0, 0, 0),
	(2, 'Daño magico debuff', 0, 'Daño magico debuff', 0, 0, 0, 0, 0, 0, 0, 0);
/*!40000 ALTER TABLE `efecto` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.feedback
CREATE TABLE IF NOT EXISTS `feedback` (
  `id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `personaje` int(255) unsigned DEFAULT NULL,
  `fecha` datetime DEFAULT NULL,
  `mensaje` text DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=318 DEFAULT CHARSET=utf8;

-- Volcando estructura para tabla aodrag.habilidad
CREATE TABLE IF NOT EXISTS `habilidad` (
  `id` smallint(5) unsigned NOT NULL AUTO_INCREMENT,
  `nombre` char(50) NOT NULL,
  `palabras_magicas` char(50) NOT NULL DEFAULT '0',
  `objetivo` tinyint(4) NOT NULL DEFAULT 0,
  `beneficiosa` tinyint(1) NOT NULL DEFAULT 0,
  `fx` mediumint(8) unsigned NOT NULL DEFAULT 0,
  `wav` mediumint(8) unsigned NOT NULL DEFAULT 0,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=6 DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.habilidad: ~5 rows (aproximadamente)
/*!40000 ALTER TABLE `habilidad` DISABLE KEYS */;
INSERT INTO `habilidad` (`id`, `nombre`, `palabras_magicas`, `objetivo`, `beneficiosa`, `fx`, `wav`) VALUES
	(1, 'Misil Magico', 'VAX IN TAR', 3, 0, 29, 111),
	(2, 'Veneno', 'VENENO Y TAL', 3, 0, 0, 0),
	(3, 'Area de fuego', 'AREA DE FUEGO', 4, 0, 0, 0),
	(4, 'Muro de hielo', 'MURO DE HIELO', 4, 0, 0, 0),
	(5, 'Curacion normal', 'CURACION', 3, 1, 0, 0);
/*!40000 ALTER TABLE `habilidad` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.mapa
CREATE TABLE IF NOT EXISTS `mapa` (
  `table_num` smallint(6) DEFAULT NULL,
  `id` smallint(6) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.mapa: ~300 rows (aproximadamente)
/*!40000 ALTER TABLE `mapa` DISABLE KEYS */;
INSERT INTO `mapa` (`table_num`, `id`) VALUES
	(1, 1),
	(2, 2),
	(3, 3),
	(4, 4),
	(5, 5),
	(6, 6),
	(7, 7),
	(8, 8),
	(9, 9),
	(10, 10),
	(11, 11),
	(12, 12),
	(13, 13),
	(14, 14),
	(15, 15),
	(16, 16),
	(17, 17),
	(18, 18),
	(19, 19),
	(20, 20),
	(21, 21),
	(22, 22),
	(23, 23),
	(24, 24),
	(25, 25),
	(26, 26),
	(27, 27),
	(28, 28),
	(29, 29),
	(30, 30),
	(31, 31),
	(32, 32),
	(33, 33),
	(34, 34),
	(35, 35),
	(36, 36),
	(37, 37),
	(38, 38),
	(39, 39),
	(40, 40),
	(41, 41),
	(42, 42),
	(43, 43),
	(44, 44),
	(45, 45),
	(46, 46),
	(47, 47),
	(48, 48),
	(49, 49),
	(50, 50),
	(51, 51),
	(52, 52),
	(53, 53),
	(54, 54),
	(55, 55),
	(56, 56),
	(57, 57),
	(58, 58),
	(59, 59),
	(60, 60),
	(61, 61),
	(62, 62),
	(63, 63),
	(64, 64),
	(65, 65),
	(66, 66),
	(67, 67),
	(68, 68),
	(69, 69),
	(70, 70),
	(71, 71),
	(72, 72),
	(73, 73),
	(74, 74),
	(75, 75),
	(76, 76),
	(77, 77),
	(78, 78),
	(79, 79),
	(80, 80),
	(81, 81),
	(82, 82),
	(83, 83),
	(84, 84),
	(85, 85),
	(86, 86),
	(87, 87),
	(88, 88),
	(89, 89),
	(90, 90),
	(91, 91),
	(92, 92),
	(93, 93),
	(94, 94),
	(95, 95),
	(96, 96),
	(97, 97),
	(98, 98),
	(99, 99),
	(100, 100),
	(101, 101),
	(102, 102),
	(103, 103),
	(104, 104),
	(105, 105),
	(106, 106),
	(107, 107),
	(108, 108),
	(109, 109),
	(110, 110),
	(111, 111),
	(112, 112),
	(113, 113),
	(114, 114),
	(115, 115),
	(116, 116),
	(117, 117),
	(118, 118),
	(119, 119),
	(120, 120),
	(121, 121),
	(122, 122),
	(123, 123),
	(124, 124),
	(125, 125),
	(126, 126),
	(127, 127),
	(128, 128),
	(129, 129),
	(130, 130),
	(131, 131),
	(132, 132),
	(133, 133),
	(134, 134),
	(135, 135),
	(136, 136),
	(137, 137),
	(138, 138),
	(139, 139),
	(140, 140),
	(141, 141),
	(142, 142),
	(143, 143),
	(144, 144),
	(145, 145),
	(146, 146),
	(147, 147),
	(148, 148),
	(149, 149),
	(150, 150),
	(151, 151),
	(152, 152),
	(153, 153),
	(154, 154),
	(155, 155),
	(156, 156),
	(157, 157),
	(158, 158),
	(159, 159),
	(160, 160),
	(161, 161),
	(162, 162),
	(163, 163),
	(164, 164),
	(165, 165),
	(166, 166),
	(167, 167),
	(168, 168),
	(169, 169),
	(170, 170),
	(171, 171),
	(172, 172),
	(173, 173),
	(174, 174),
	(175, 175),
	(176, 176),
	(177, 177),
	(178, 178),
	(179, 179),
	(180, 180),
	(181, 181),
	(182, 182),
	(183, 183),
	(184, 184),
	(185, 185),
	(186, 186),
	(187, 187),
	(188, 188),
	(189, 189),
	(190, 190),
	(191, 191),
	(192, 192),
	(193, 193),
	(194, 194),
	(195, 195),
	(196, 196),
	(197, 197),
	(198, 198),
	(199, 199),
	(200, 200),
	(201, 201),
	(202, 202),
	(203, 203),
	(204, 204),
	(205, 205),
	(206, 206),
	(207, 207),
	(208, 208),
	(209, 209),
	(210, 210),
	(211, 211),
	(212, 212),
	(213, 213),
	(214, 214),
	(215, 215),
	(216, 216),
	(217, 217),
	(218, 218),
	(219, 219),
	(220, 220),
	(221, 221),
	(222, 222),
	(223, 223),
	(224, 224),
	(225, 225),
	(226, 226),
	(227, 227),
	(228, 228),
	(229, 229),
	(230, 230),
	(231, 231),
	(232, 232),
	(233, 233),
	(234, 234),
	(235, 235),
	(236, 236),
	(237, 237),
	(238, 238),
	(239, 239),
	(240, 240),
	(241, 241),
	(242, 242),
	(243, 243),
	(244, 244),
	(245, 245),
	(246, 246),
	(247, 247),
	(248, 248),
	(249, 249),
	(250, 250),
	(251, 251),
	(252, 252),
	(253, 253),
	(254, 254),
	(255, 255),
	(256, 256),
	(257, 257),
	(258, 258),
	(259, 259),
	(260, 260),
	(261, 261),
	(262, 262),
	(263, 263),
	(264, 264),
	(265, 265),
	(266, 266),
	(267, 267),
	(268, 268),
	(269, 269),
	(270, 270),
	(271, 271),
	(272, 272),
	(273, 273),
	(274, 274),
	(275, 275),
	(276, 276),
	(277, 277),
	(278, 278),
	(279, 279),
	(280, 280),
	(281, 281),
	(282, 282),
	(283, 283),
	(284, 284),
	(285, 285),
	(286, 286),
	(287, 287),
	(288, 288),
	(289, 289),
	(290, 290),
	(291, 291),
	(292, 292),
	(293, 293),
	(294, 294),
	(295, 295),
	(296, 296),
	(297, 297),
	(298, 298),
	(299, 299),
	(300, 300);
/*!40000 ALTER TABLE `mapa` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.montura
CREATE TABLE IF NOT EXISTS `montura` (
  `id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `id_personaje` int(10) unsigned DEFAULT NULL,
  `tipo` varchar(50) DEFAULT NULL,
  `nombre` varchar(50) DEFAULT NULL,
  `level` tinyint(255) unsigned DEFAULT NULL,
  `skills` varchar(255) DEFAULT NULL,
  `elu` smallint(10) DEFAULT NULL,
  `exp` smallint(10) DEFAULT NULL,
  `ataque` smallint(10) DEFAULT NULL,
  `defensa` smallint(10) DEFAULT NULL,
  `atmagia` smallint(10) DEFAULT NULL,
  `defmagia` smallint(10) DEFAULT NULL,
  `evasion` smallint(10) DEFAULT NULL,
  `speed` smallint(6) DEFAULT 0,
  `fecha_creacion` datetime DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.montura: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `montura` DISABLE KEYS */;
INSERT INTO `montura` (`id`, `id_personaje`, `tipo`, `nombre`, `level`, `skills`, `elu`, `exp`, `ataque`, `defensa`, `atmagia`, `defmagia`, `evasion`, `speed`, `fecha_creacion`) VALUES
	(1, 3, '2', 'Dragón Rojo', 1, '0', 30, 0, 0, 0, 0, 0, 0, 5, '2018-11-21 15:20:43');
/*!40000 ALTER TABLE `montura` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.npc
CREATE TABLE IF NOT EXISTS `npc` (
  `id` int(11) DEFAULT NULL,
  `name` char(50) DEFAULT NULL,
  `desc` char(50) DEFAULT NULL,
  `nivel` tinyint(4) DEFAULT NULL,
  `movement` tinyint(4) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.npc: ~2 rows (aproximadamente)
/*!40000 ALTER TABLE `npc` DISABLE KEYS */;
INSERT INTO `npc` (`id`, `name`, `desc`, `nivel`, `movement`) VALUES
	(645, 'NPC645', NULL, NULL, NULL),
	(646, 'NPC646', NULL, NULL, NULL);
/*!40000 ALTER TABLE `npc` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.party
CREATE TABLE IF NOT EXISTS `party` (
  `id` bigint(20) unsigned NOT NULL AUTO_INCREMENT,
  `lider` int(255) unsigned DEFAULT NULL,
  `fecha_creacion` datetime DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.party: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `party` DISABLE KEYS */;
/*!40000 ALTER TABLE `party` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.personaje
CREATE TABLE IF NOT EXISTS `personaje` (
  `id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `id_cuenta` int(10) unsigned DEFAULT NULL,
  `id_clan` smallint(255) unsigned DEFAULT 0,
  `nombre` varchar(20) NOT NULL,
  `elv` tinyint(255) unsigned DEFAULT 1,
  `logged` tinyint(255) DEFAULT 0,
  `genero` tinyint(255) DEFAULT NULL,
  `raza` tinyint(255) DEFAULT NULL,
  `clase` tinyint(255) DEFAULT NULL,
  `heading` tinyint(255) DEFAULT 3,
  `head` smallint(255) DEFAULT NULL,
  `body` smallint(255) DEFAULT 0,
  `arma` smallint(255) DEFAULT 0,
  `escudo` smallint(255) DEFAULT 0,
  `casco` smallint(255) DEFAULT 0,
  `uptime` mediumint(8) unsigned DEFAULT 0,
  `lastip1` varchar(50) DEFAULT '',
  `position` varchar(10) DEFAULT '1-40-86',
  `descripcion` varchar(150) DEFAULT '',
  `muerto` tinyint(255) unsigned DEFAULT NULL,
  `escondido` tinyint(255) unsigned DEFAULT NULL,
  `advertencias` smallint(255) unsigned DEFAULT NULL,
  `hambre` tinyint(255) unsigned DEFAULT NULL,
  `sed` tinyint(255) unsigned DEFAULT NULL,
  `desnudo` tinyint(255) unsigned DEFAULT NULL,
  `ban` tinyint(255) unsigned DEFAULT NULL,
  `navegando` tinyint(255) unsigned DEFAULT NULL,
  `quemontura` tinyint(255) unsigned DEFAULT NULL,
  `envenenado` tinyint(255) unsigned DEFAULT NULL,
  `inmovilizado` tinyint(255) unsigned DEFAULT NULL,
  `paralizado` tinyint(255) unsigned DEFAULT NULL,
  `serialhd` varchar(25) DEFAULT NULL,
  `pertenece` tinyint(255) unsigned DEFAULT NULL,
  `pertenececaos` tinyint(255) unsigned DEFAULT NULL,
  `pena` smallint(255) unsigned DEFAULT NULL,
  `ejercitoreal` tinyint(255) unsigned DEFAULT NULL,
  `ejercitocaos` tinyint(255) unsigned DEFAULT NULL,
  `ciudmatados` int(255) unsigned DEFAULT NULL,
  `crimmatados` int(255) unsigned DEFAULT NULL,
  `rarcaos` tinyint(255) unsigned DEFAULT NULL,
  `rarreal` tinyint(255) unsigned DEFAULT NULL,
  `rexcaos` tinyint(255) unsigned DEFAULT NULL,
  `rexreal` tinyint(255) unsigned DEFAULT NULL,
  `reccaos` tinyint(255) unsigned DEFAULT NULL,
  `recreal` tinyint(255) unsigned DEFAULT NULL,
  `reenlistadas` tinyint(255) unsigned DEFAULT NULL,
  `nivelingreso` tinyint(255) unsigned DEFAULT NULL,
  `fechaingreso` varchar(50) DEFAULT NULL,
  `matadosingreso` smallint(255) unsigned DEFAULT NULL,
  `nextrecompensa` smallint(255) unsigned DEFAULT NULL,
  `gld` int(255) unsigned DEFAULT NULL,
  `banco` int(255) unsigned DEFAULT NULL,
  `maxhp` smallint(255) unsigned DEFAULT NULL,
  `minhp` smallint(255) unsigned DEFAULT NULL,
  `maxsta` smallint(255) unsigned DEFAULT NULL,
  `minsta` smallint(255) unsigned DEFAULT NULL,
  `maxman` smallint(255) unsigned DEFAULT NULL,
  `minman` smallint(255) unsigned DEFAULT NULL,
  `maxhit` smallint(255) unsigned DEFAULT NULL,
  `minhit` smallint(255) unsigned DEFAULT NULL,
  `maxagu` tinyint(255) unsigned DEFAULT NULL,
  `minagu` tinyint(255) unsigned DEFAULT NULL,
  `maxham` tinyint(255) unsigned DEFAULT NULL,
  `minham` tinyint(255) unsigned DEFAULT NULL,
  `exp` int(255) unsigned DEFAULT NULL,
  `elu` int(10) unsigned DEFAULT NULL,
  `elo` smallint(255) unsigned DEFAULT NULL,
  `dragcredits` smallint(255) unsigned DEFAULT NULL,
  `nummonturas` tinyint(255) unsigned DEFAULT NULL,
  `usermuertes` int(255) unsigned DEFAULT NULL,
  `npcsmuertes` int(255) unsigned DEFAULT NULL,
  `cantidaditems` tinyint(10) DEFAULT NULL,
  `weaponeqpslot` tinyint(255) DEFAULT NULL,
  `armoureqpslot` tinyint(255) DEFAULT NULL,
  `cascoeqpslot` tinyint(255) DEFAULT NULL,
  `escudoeqpslot` tinyint(255) DEFAULT NULL,
  `barcoslot` int(255) DEFAULT NULL,
  `municionslot` tinyint(255) DEFAULT NULL,
  `anilloslot` varchar(255) DEFAULT NULL,
  `asesino` int(11) DEFAULT NULL,
  `bandido` int(255) DEFAULT NULL,
  `burguesia` int(255) DEFAULT NULL,
  `ladrones` int(255) DEFAULT NULL,
  `nobles` int(255) DEFAULT NULL,
  `plebe` int(255) DEFAULT NULL,
  `promedio` int(255) unsigned DEFAULT NULL,
  `seguro` tinyint(1) DEFAULT NULL,
  `userindex` smallint(255) unsigned DEFAULT NULL,
  `borrado` tinyint(3) unsigned DEFAULT NULL,
  PRIMARY KEY (`nombre`),
  KEY `id` (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=13 DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.personaje: ~8 rows (aproximadamente)
/*!40000 ALTER TABLE `personaje` DISABLE KEYS */;
INSERT INTO `personaje` (`id`, `id_cuenta`, `id_clan`, `nombre`, `elv`, `logged`, `genero`, `raza`, `clase`, `heading`, `head`, `body`, `arma`, `escudo`, `casco`, `uptime`, `lastip1`, `position`, `descripcion`, `muerto`, `escondido`, `advertencias`, `hambre`, `sed`, `desnudo`, `ban`, `navegando`, `quemontura`, `envenenado`, `inmovilizado`, `paralizado`, `serialhd`, `pertenece`, `pertenececaos`, `pena`, `ejercitoreal`, `ejercitocaos`, `ciudmatados`, `crimmatados`, `rarcaos`, `rarreal`, `rexcaos`, `rexreal`, `reccaos`, `recreal`, `reenlistadas`, `nivelingreso`, `fechaingreso`, `matadosingreso`, `nextrecompensa`, `gld`, `banco`, `maxhp`, `minhp`, `maxsta`, `minsta`, `maxman`, `minman`, `maxhit`, `minhit`, `maxagu`, `minagu`, `maxham`, `minham`, `exp`, `elu`, `elo`, `dragcredits`, `nummonturas`, `usermuertes`, `npcsmuertes`, `cantidaditems`, `weaponeqpslot`, `armoureqpslot`, `cascoeqpslot`, `escudoeqpslot`, `barcoslot`, `municionslot`, `anilloslot`, `asesino`, `bandido`, `burguesia`, `ladrones`, `nobles`, `plebe`, `promedio`, `seguro`, `userindex`, `borrado`) VALUES
	(2, 1, 0, 'AODrag', 58, 0, 1, 2, 2, 1, 101, 348, 2, 2, 2, 0, '16777343', '166-45-81', 'hola amigos', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '282975812', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 'No ingresó a ninguna Facción', 0, 0, 11880, 0, 500, 500, 205, 205, 3453, 1093, 12, 11, 100, 100, 100, 100, 567, 1224, 1000, 0, 0, 0, 159, 80, 0, 11, 0, 0, 9, 0, '0', 0, 0, 0, 0, 80500, 30, 13422, 0, 5, 0),
	(11, 1, 0, 'asdf', 50, 0, 2, 1, 2, 1, 70, 226, 6, 2, 2, 0, '16777343', '166-39-162', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '1959897935', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 'No ingresó a ninguna Facción', 0, 0, 50, 0, 500, 500, 50, 50, 50, 50, 2, 1, 100, 74, 100, 80, 292, 300, 1000, 0, 0, 0, 3, 80, 6, 5, 0, 0, 0, 0, '0', 0, 0, 0, 0, 2500, 30, 422, 0, 11, 0),
	(12, 1, 0, 'Irongete', 6, 0, 1, 1, 1, 1, 1, 226, 6, 2, 2, 0, '16777343', '166-38-176', '', 0, 0, 0, 1, 1, 0, 0, 0, 0, 0, 0, 0, '1959897935', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 'No ingresó a ninguna Facción', 0, 0, 375, 0, 66, 61, 185, 154, 302, 63, 7, 6, 100, 0, 100, 0, 72, 483, 1000, 0, 0, 0, 25, 80, 6, 4, 0, 0, 0, 0, '0', 0, 0, 0, 0, 13500, 30, 2255, 0, 2, 0),
	(3, 2, 0, 'Lorwik', 100, 1, 1, 1, 2, 1, 4, 318, 0, 5, 2, 0, '16777343', '166-32-163', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '282975812', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 'No ingresó a ninguna Facción', 0, 0, 9985815, 0, 500, 500, 2034, 2034, 2406, 54, 126, 125, 100, 90, 100, 90, 77409401, 1556129566, 1000, 0, 1, 0, 59, 80, 12, 15, 0, 21, 24, 0, '0', 0, 0, 0, 0, 30500, 30, 5088, 0, 2, 0),
	(1, 1, 0, 'magohumano', 54, 0, 1, 1, 1, 1, 1, 174, 2, 2, 2, 0, '16777343', '166-56-47', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '1959897935', 0, 0, 0, 1, 0, 0, 0, 0, 1, 0, 0, 0, 0, 1, 54, '06/12/2018', 0, 30, 3796, 5, 500, 500, 131, 131, 204, 84, 5, 4, 100, 30, 100, 50, 132, 1012, 1000, 0, 0, 0, 65, 80, 0, 1, 0, 0, 0, 0, '0', 0, 0, 0, 0, 33500, 30, 5588, 0, 8, 0),
	(9, 2, 0, 'Orco', 50, 0, 1, 6, 2, 1, 130, 24, 2, 2, 2, 0, '16777343', '166-45-102', '', 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, '1959897935', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 'No ingresó a ninguna Facción', 0, 0, 0, 0, 500, 500, 50, 0, 50, 50, 2, 1, 100, 90, 100, 90, 5, 300, 1000, 0, 0, 0, 0, 80, 0, 0, 0, 0, 0, 0, '0', 0, 0, 0, 0, 1000, 30, 172, 0, 2, 0),
	(6, 2, 0, 'PENE', 3, 0, 1, 1, 2, 1, 1, 226, 6, 2, 2, 0, '16777343', '166-45-148', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '282975812', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 'No ingresó a ninguna Facción', 0, 0, 140, 0, 500, 500, 114, 114, 126, 106, 6, 5, 100, 100, 100, 100, 94, 363, 1000, 0, 0, 0, 11, 80, 6, 5, 0, 0, 0, 0, '0', 0, 0, 0, 0, 6500, 30, 1088, 0, 7, 0),
	(10, 2, 0, 'Pepe', 1, 0, 1, 1, 2, 1, 5, 226, 2, 2, 2, 0, '16777343', '1-44-88', '', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '282975812', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 'No ingresó a ninguna Facción', 0, 0, 0, 0, 500, 500, 50, 50, 50, 50, 2, 1, 100, 100, 100, 100, 0, 300, 1000, 0, 0, 0, 0, 80, 0, 4, 0, 0, 0, 0, '0', 0, 0, 0, 0, 1000, 30, 172, 0, 5, 0);
/*!40000 ALTER TABLE `personaje` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.quest
CREATE TABLE IF NOT EXISTS `quest` (
  `id` smallint(6) unsigned NOT NULL,
  `id_prequest` smallint(6) unsigned DEFAULT NULL,
  `nivel_minimo` tinyint(3) unsigned NOT NULL,
  `titulo` char(50) NOT NULL,
  `texto` text NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.quest: ~3 rows (aproximadamente)
/*!40000 ALTER TABLE `quest` DISABLE KEYS */;
INSERT INTO `quest` (`id`, `id_prequest`, `nivel_minimo`, `titulo`, `texto`) VALUES
	(1, NULL, 1, 'Primera quest de prueba', 'Mata un lobo.'),
	(2, 1, 1, 'Segunda quest de prueba', 'Mataste un lobo, mata otro lobo.'),
	(3, NULL, 1, 'Tercera quest de prueba', 'Mata una serpiente.');
/*!40000 ALTER TABLE `quest` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_clan_puntos
CREATE TABLE IF NOT EXISTS `rel_clan_puntos` (
  `id_clan` int(10) unsigned NOT NULL,
  `puntoscastillos` mediumint(10) unsigned DEFAULT NULL,
  `minutoscastillo1` smallint(255) unsigned DEFAULT NULL,
  `minutoscastillo2` smallint(255) unsigned DEFAULT NULL,
  `minutoscastillo3` smallint(255) unsigned DEFAULT NULL,
  `minutoscastillo4` smallint(255) unsigned DEFAULT NULL,
  `minutoscastillo5` smallint(255) unsigned DEFAULT NULL,
  PRIMARY KEY (`id_clan`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_clan_puntos: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_clan_puntos` DISABLE KEYS */;
/*!40000 ALTER TABLE `rel_clan_puntos` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_clan_solicitud
CREATE TABLE IF NOT EXISTS `rel_clan_solicitud` (
  `id_clan` smallint(255) unsigned DEFAULT NULL,
  `id_personaje` int(255) unsigned DEFAULT NULL,
  `fecha` datetime DEFAULT NULL,
  `mensaje` varchar(250) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_clan_solicitud: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_clan_solicitud` DISABLE KEYS */;
/*!40000 ALTER TABLE `rel_clan_solicitud` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_cuenta_boveda
CREATE TABLE IF NOT EXISTS `rel_cuenta_boveda` (
  `id_cuenta` int(255) unsigned NOT NULL,
  `slot` smallint(255) unsigned NOT NULL,
  `objeto` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`id_cuenta`,`slot`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_cuenta_boveda: ~160 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_cuenta_boveda` DISABLE KEYS */;
INSERT INTO `rel_cuenta_boveda` (`id_cuenta`, `slot`, `objeto`) VALUES
	(1, 1, '0-0'),
	(1, 2, '0-0'),
	(1, 3, '0-0'),
	(1, 4, '0-0'),
	(1, 5, '0-0'),
	(1, 6, '0-0'),
	(1, 7, '0-0'),
	(1, 8, '0-0'),
	(1, 9, '0-0'),
	(1, 10, '0-0'),
	(1, 11, '0-0'),
	(1, 12, '0-0'),
	(1, 13, '0-0'),
	(1, 14, '0-0'),
	(1, 15, '0-0'),
	(1, 16, '0-0'),
	(1, 17, '0-0'),
	(1, 18, '0-0'),
	(1, 19, '0-0'),
	(1, 20, '0-0'),
	(1, 21, '0-0'),
	(1, 22, '0-0'),
	(1, 23, '0-0'),
	(1, 24, '0-0'),
	(1, 25, '0-0'),
	(1, 26, '0-0'),
	(1, 27, '0-0'),
	(1, 28, '0-0'),
	(1, 29, '0-0'),
	(1, 30, '0-0'),
	(1, 31, '0-0'),
	(1, 32, '0-0'),
	(1, 33, '0-0'),
	(1, 34, '0-0'),
	(1, 35, '0-0'),
	(1, 36, '0-0'),
	(1, 37, '0-0'),
	(1, 38, '0-0'),
	(1, 39, '0-0'),
	(1, 40, '0-0'),
	(1, 41, '0-0'),
	(1, 42, '0-0'),
	(1, 43, '0-0'),
	(1, 44, '0-0'),
	(1, 45, '0-0'),
	(1, 46, '0-0'),
	(1, 47, '0-0'),
	(1, 48, '0-0'),
	(1, 49, '0-0'),
	(1, 50, '0-0'),
	(1, 51, '0-0'),
	(1, 52, '0-0'),
	(1, 53, '0-0'),
	(1, 54, '0-0'),
	(1, 55, '0-0'),
	(1, 56, '0-0'),
	(1, 57, '0-0'),
	(1, 58, '0-0'),
	(1, 59, '0-0'),
	(1, 60, '0-0'),
	(1, 61, '0-0'),
	(1, 62, '0-0'),
	(1, 63, '0-0'),
	(1, 64, '0-0'),
	(1, 65, '0-0'),
	(1, 66, '0-0'),
	(1, 67, '0-0'),
	(1, 68, '0-0'),
	(1, 69, '0-0'),
	(1, 70, '0-0'),
	(1, 71, '0-0'),
	(1, 72, '0-0'),
	(1, 73, '0-0'),
	(1, 74, '0-0'),
	(1, 75, '0-0'),
	(1, 76, '0-0'),
	(1, 77, '0-0'),
	(1, 78, '0-0'),
	(1, 79, '0-0'),
	(1, 80, '0-0'),
	(2, 1, '0-0'),
	(2, 2, '0-0'),
	(2, 3, '0-0'),
	(2, 4, '0-0'),
	(2, 5, '0-0'),
	(2, 6, '0-0'),
	(2, 7, '0-0'),
	(2, 8, '0-0'),
	(2, 9, '0-0'),
	(2, 10, '0-0'),
	(2, 11, '0-0'),
	(2, 12, '0-0'),
	(2, 13, '0-0'),
	(2, 14, '0-0'),
	(2, 15, '0-0'),
	(2, 16, '0-0'),
	(2, 17, '0-0'),
	(2, 18, '0-0'),
	(2, 19, '0-0'),
	(2, 20, '0-0'),
	(2, 21, '0-0'),
	(2, 22, '0-0'),
	(2, 23, '0-0'),
	(2, 24, '0-0'),
	(2, 25, '0-0'),
	(2, 26, '0-0'),
	(2, 27, '0-0'),
	(2, 28, '0-0'),
	(2, 29, '0-0'),
	(2, 30, '0-0'),
	(2, 31, '0-0'),
	(2, 32, '0-0'),
	(2, 33, '0-0'),
	(2, 34, '0-0'),
	(2, 35, '0-0'),
	(2, 36, '0-0'),
	(2, 37, '0-0'),
	(2, 38, '0-0'),
	(2, 39, '0-0'),
	(2, 40, '0-0'),
	(2, 41, '0-0'),
	(2, 42, '0-0'),
	(2, 43, '0-0'),
	(2, 44, '0-0'),
	(2, 45, '0-0'),
	(2, 46, '0-0'),
	(2, 47, '0-0'),
	(2, 48, '0-0'),
	(2, 49, '0-0'),
	(2, 50, '0-0'),
	(2, 51, '0-0'),
	(2, 52, '0-0'),
	(2, 53, '0-0'),
	(2, 54, '0-0'),
	(2, 55, '0-0'),
	(2, 56, '0-0'),
	(2, 57, '0-0'),
	(2, 58, '0-0'),
	(2, 59, '0-0'),
	(2, 60, '0-0'),
	(2, 61, '0-0'),
	(2, 62, '0-0'),
	(2, 63, '0-0'),
	(2, 64, '0-0'),
	(2, 65, '0-0'),
	(2, 66, '0-0'),
	(2, 67, '0-0'),
	(2, 68, '0-0'),
	(2, 69, '0-0'),
	(2, 70, '0-0'),
	(2, 71, '0-0'),
	(2, 72, '0-0'),
	(2, 73, '0-0'),
	(2, 74, '0-0'),
	(2, 75, '0-0'),
	(2, 76, '0-0'),
	(2, 77, '0-0'),
	(2, 78, '0-0'),
	(2, 79, '0-0'),
	(2, 80, '0-0');
/*!40000 ALTER TABLE `rel_cuenta_boveda` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_cuenta_desarrollo
CREATE TABLE IF NOT EXISTS `rel_cuenta_desarrollo` (
  `id_cuenta` int(255) unsigned DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_cuenta_desarrollo: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_cuenta_desarrollo` DISABLE KEYS */;
/*!40000 ALTER TABLE `rel_cuenta_desarrollo` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_efecto_trigger
CREATE TABLE IF NOT EXISTS `rel_efecto_trigger` (
  `id_efecto` smallint(5) unsigned NOT NULL,
  `trigger` smallint(5) unsigned NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_efecto_trigger: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_efecto_trigger` DISABLE KEYS */;
/*!40000 ALTER TABLE `rel_efecto_trigger` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_efecto_valor
CREATE TABLE IF NOT EXISTS `rel_efecto_valor` (
  `id_efecto` int(11) DEFAULT NULL,
  `tipo` tinyint(4) DEFAULT NULL,
  `valor` mediumint(8) unsigned DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_efecto_valor: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_efecto_valor` DISABLE KEYS */;
/*!40000 ALTER TABLE `rel_efecto_valor` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_habilidad_efecto
CREATE TABLE IF NOT EXISTS `rel_habilidad_efecto` (
  `id_habilidad` smallint(6) NOT NULL,
  `id_efecto` smallint(6) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_habilidad_efecto: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_habilidad_efecto` DISABLE KEYS */;
/*!40000 ALTER TABLE `rel_habilidad_efecto` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_party_invitacion
CREATE TABLE IF NOT EXISTS `rel_party_invitacion` (
  `id_invita` int(255) unsigned DEFAULT NULL,
  `id_invitado` int(255) unsigned DEFAULT NULL,
  `fecha` datetime DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- Volcando datos para la tabla aodrag.rel_party_invitacion: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_party_invitacion` DISABLE KEYS */;
/*!40000 ALTER TABLE `rel_party_invitacion` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_party_personaje
CREATE TABLE IF NOT EXISTS `rel_party_personaje` (
  `id_party` int(255) unsigned DEFAULT NULL,
  `id_personaje` int(255) unsigned DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- Volcando datos para la tabla aodrag.rel_party_personaje: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_party_personaje` DISABLE KEYS */;
/*!40000 ALTER TABLE `rel_party_personaje` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_personaje_atributo
CREATE TABLE IF NOT EXISTS `rel_personaje_atributo` (
  `id_personaje` int(255) unsigned DEFAULT NULL,
  `atributo` mediumint(255) unsigned DEFAULT NULL,
  `valor` smallint(255) unsigned DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_personaje_atributo: ~35 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_personaje_atributo` DISABLE KEYS */;
INSERT INTO `rel_personaje_atributo` (`id_personaje`, `atributo`, `valor`) VALUES
	(1, 1, 17),
	(1, 2, 18),
	(1, 3, 19),
	(1, 4, 18),
	(1, 5, 20),
	(2, 1, 17),
	(2, 2, 18),
	(2, 3, 20),
	(2, 4, 17),
	(2, 5, 19),
	(3, 1, 18),
	(3, 2, 17),
	(3, 3, 19),
	(3, 4, 18),
	(3, 5, 20),
	(6, 1, 18),
	(6, 2, 17),
	(6, 3, 19),
	(6, 4, 18),
	(6, 5, 20),
	(9, 1, 21),
	(9, 2, 14),
	(9, 3, 13),
	(9, 4, 16),
	(9, 5, 21),
	(10, 1, 18),
	(10, 2, 17),
	(10, 3, 19),
	(10, 4, 18),
	(10, 5, 20),
	(11, 1, 18),
	(11, 2, 17),
	(11, 3, 19),
	(11, 4, 18),
	(11, 5, 20),
	(12, 1, 17),
	(12, 2, 18),
	(12, 3, 19),
	(12, 4, 18),
	(12, 5, 20);
/*!40000 ALTER TABLE `rel_personaje_atributo` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_personaje_habilidad
CREATE TABLE IF NOT EXISTS `rel_personaje_habilidad` (
  `id_personaje` mediumint(9) NOT NULL,
  `slot` tinyint(3) unsigned NOT NULL,
  `id_habilidad` smallint(6) NOT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_personaje_habilidad: ~5 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_personaje_habilidad` DISABLE KEYS */;
INSERT INTO `rel_personaje_habilidad` (`id_personaje`, `slot`, `id_habilidad`) VALUES
	(12, 1, 1),
	(12, 2, 2),
	(12, 3, 3),
	(12, 4, 4),
	(12, 5, 5);
/*!40000 ALTER TABLE `rel_personaje_habilidad` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_personaje_hechizo
CREATE TABLE IF NOT EXISTS `rel_personaje_hechizo` (
  `id_personaje` int(255) unsigned NOT NULL,
  `slot` tinyint(255) NOT NULL,
  `hechizo` smallint(255) unsigned DEFAULT NULL,
  PRIMARY KEY (`id_personaje`,`slot`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_personaje_hechizo: ~144 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_personaje_hechizo` DISABLE KEYS */;
INSERT INTO `rel_personaje_hechizo` (`id_personaje`, `slot`, `hechizo`) VALUES
	(1, 1, 2),
	(1, 2, 44),
	(1, 3, 14),
	(1, 4, 0),
	(1, 5, 0),
	(1, 6, 0),
	(1, 7, 0),
	(1, 8, 0),
	(1, 9, 0),
	(1, 10, 0),
	(1, 11, 0),
	(1, 12, 0),
	(1, 13, 0),
	(1, 14, 0),
	(1, 15, 0),
	(1, 16, 0),
	(1, 17, 0),
	(1, 18, 0),
	(2, 1, 0),
	(2, 2, 0),
	(2, 3, 0),
	(2, 4, 0),
	(2, 5, 0),
	(2, 6, 0),
	(2, 7, 0),
	(2, 8, 0),
	(2, 9, 0),
	(2, 10, 0),
	(2, 11, 0),
	(2, 12, 0),
	(2, 13, 0),
	(2, 14, 0),
	(2, 15, 0),
	(2, 16, 0),
	(2, 17, 0),
	(2, 18, 0),
	(3, 1, 0),
	(3, 2, 0),
	(3, 3, 0),
	(3, 4, 0),
	(3, 5, 0),
	(3, 6, 0),
	(3, 7, 0),
	(3, 8, 0),
	(3, 9, 0),
	(3, 10, 0),
	(3, 11, 0),
	(3, 12, 0),
	(3, 13, 0),
	(3, 14, 0),
	(3, 15, 0),
	(3, 16, 0),
	(3, 17, 0),
	(3, 18, 0),
	(6, 1, 0),
	(6, 2, 0),
	(6, 3, 0),
	(6, 4, 0),
	(6, 5, 0),
	(6, 6, 0),
	(6, 7, 0),
	(6, 8, 0),
	(6, 9, 0),
	(6, 10, 0),
	(6, 11, 0),
	(6, 12, 0),
	(6, 13, 0),
	(6, 14, 0),
	(6, 15, 0),
	(6, 16, 0),
	(6, 17, 0),
	(6, 18, 0),
	(9, 1, 0),
	(9, 2, 0),
	(9, 3, 0),
	(9, 4, 0),
	(9, 5, 0),
	(9, 6, 0),
	(9, 7, 0),
	(9, 8, 0),
	(9, 9, 0),
	(9, 10, 0),
	(9, 11, 0),
	(9, 12, 0),
	(9, 13, 0),
	(9, 14, 0),
	(9, 15, 0),
	(9, 16, 0),
	(9, 17, 0),
	(9, 18, 0),
	(10, 1, 2),
	(10, 2, 1),
	(10, 3, 0),
	(10, 4, 0),
	(10, 5, 0),
	(10, 6, 0),
	(10, 7, 0),
	(10, 8, 0),
	(10, 9, 0),
	(10, 10, 0),
	(10, 11, 0),
	(10, 12, 0),
	(10, 13, 0),
	(10, 14, 0),
	(10, 15, 0),
	(10, 16, 0),
	(10, 17, 0),
	(10, 18, 0),
	(11, 1, 2),
	(11, 2, 1),
	(11, 3, 0),
	(11, 4, 0),
	(11, 5, 0),
	(11, 6, 0),
	(11, 7, 0),
	(11, 8, 0),
	(11, 9, 0),
	(11, 10, 0),
	(11, 11, 0),
	(11, 12, 0),
	(11, 13, 0),
	(11, 14, 0),
	(11, 15, 0),
	(11, 16, 0),
	(11, 17, 0),
	(11, 18, 0),
	(12, 1, 1),
	(12, 2, 2),
	(12, 3, 3),
	(12, 4, 4),
	(12, 5, 5),
	(12, 6, 0),
	(12, 7, 0),
	(12, 8, 0),
	(12, 9, 0),
	(12, 10, 0),
	(12, 11, 0),
	(12, 12, 0),
	(12, 13, 0),
	(12, 14, 0),
	(12, 15, 0),
	(12, 16, 0),
	(12, 17, 0),
	(12, 18, 0);
/*!40000 ALTER TABLE `rel_personaje_hechizo` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_personaje_inventario
CREATE TABLE IF NOT EXISTS `rel_personaje_inventario` (
  `id_personaje` int(255) unsigned NOT NULL,
  `slot` int(255) unsigned NOT NULL,
  `objeto` varchar(20) DEFAULT NULL,
  PRIMARY KEY (`id_personaje`,`slot`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_personaje_inventario: ~175 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_personaje_inventario` DISABLE KEYS */;
INSERT INTO `rel_personaje_inventario` (`id_personaje`, `slot`, `objeto`) VALUES
	(1, 1, '675-1-1'),
	(1, 2, '208-1-0'),
	(1, 3, '0-0-0'),
	(1, 4, '0-0-0'),
	(1, 5, '0-0-0'),
	(1, 6, '0-0-0'),
	(1, 7, '0-0-0'),
	(1, 8, '0-0-0'),
	(1, 9, '0-0-0'),
	(1, 10, '0-0-0'),
	(1, 11, '0-0-0'),
	(1, 12, '0-0-0'),
	(1, 13, '0-0-0'),
	(1, 14, '0-0-0'),
	(1, 15, '0-0-0'),
	(1, 16, '0-0-0'),
	(1, 17, '0-0-0'),
	(1, 18, '0-0-0'),
	(1, 19, '0-0-0'),
	(1, 20, '0-0-0'),
	(1, 21, '0-0-0'),
	(2, 1, '467-98-0'),
	(2, 2, '468-99-0'),
	(2, 3, '461-50-0'),
	(2, 4, '127-1-0'),
	(2, 5, '463-1-1'),
	(2, 6, '1020-1-0'),
	(2, 7, '465-100-0'),
	(2, 8, '474-1-0'),
	(2, 9, '475-1-0'),
	(2, 10, '476-1-0'),
	(2, 11, '952-100-1'),
	(2, 12, '0-0-0'),
	(2, 13, '0-0-0'),
	(2, 14, '0-0-0'),
	(2, 15, '58-7-0'),
	(2, 16, '164-1-1'),
	(2, 17, '15-1-0'),
	(2, 18, '0-0-0'),
	(2, 19, '0-0-0'),
	(2, 20, '0-0-0'),
	(2, 21, '0-0-0'),
	(3, 1, '123-100-0'),
	(3, 2, '30-1-0'),
	(3, 3, '187-100-0'),
	(3, 4, '386-46-0'),
	(3, 5, '388-495-0'),
	(3, 6, '1027-100-0'),
	(3, 7, '1028-99-0'),
	(3, 8, '1100-100-0'),
	(3, 9, '389-100-0'),
	(3, 10, '1094-100-0'),
	(3, 11, '387-113-0'),
	(3, 12, '543-1-1'),
	(3, 13, '132-1-0'),
	(3, 14, '192-21-0'),
	(3, 15, '905-100-1'),
	(3, 16, '193-2-0'),
	(3, 17, '139-10-0'),
	(3, 18, '478-99-0'),
	(3, 19, '164-1-0'),
	(3, 20, '0-0-0'),
	(3, 21, '708-100-1'),
	(6, 1, '467-100-0'),
	(6, 2, '468-96-0'),
	(6, 3, '461-16-0'),
	(6, 4, '551-90-0'),
	(6, 5, '463-1-1'),
	(6, 6, '1020-1-1'),
	(6, 7, '465-73-0'),
	(6, 8, '478-1-0'),
	(6, 9, '0-0-0'),
	(6, 10, '0-0-0'),
	(6, 11, '0-0-0'),
	(6, 12, '0-0-0'),
	(6, 13, '0-0-0'),
	(6, 14, '0-0-0'),
	(6, 15, '0-0-0'),
	(6, 16, '0-0-0'),
	(6, 17, '0-0-0'),
	(6, 18, '0-0-0'),
	(6, 19, '0-0-0'),
	(6, 20, '0-0-0'),
	(6, 21, '0-0-0'),
	(9, 1, '467-100-0'),
	(9, 2, '468-100-0'),
	(9, 3, '461-50-0'),
	(9, 4, '0-0-0'),
	(9, 5, '463-1-0'),
	(9, 6, '1020-1-0'),
	(9, 7, '465-100-0'),
	(9, 8, '0-0-0'),
	(9, 9, '0-0-0'),
	(9, 10, '0-0-0'),
	(9, 11, '0-0-0'),
	(9, 12, '0-0-0'),
	(9, 13, '0-0-0'),
	(9, 14, '0-0-0'),
	(9, 15, '0-0-0'),
	(9, 16, '0-0-0'),
	(9, 17, '0-0-0'),
	(9, 18, '0-0-0'),
	(9, 19, '0-0-0'),
	(9, 20, '0-0-0'),
	(9, 21, '0-0-0'),
	(10, 1, '467-100-0'),
	(10, 2, '468-100-0'),
	(10, 3, '461-50-0'),
	(10, 4, '0-0-0'),
	(10, 5, '463-1-1'),
	(10, 6, '1020-1-0'),
	(10, 7, '465-100-0'),
	(10, 8, '0-0-0'),
	(10, 9, '0-0-0'),
	(10, 10, '0-0-0'),
	(10, 11, '0-0-0'),
	(10, 12, '0-0-0'),
	(10, 13, '0-0-0'),
	(10, 14, '0-0-0'),
	(10, 15, '0-0-0'),
	(10, 16, '0-0-0'),
	(10, 17, '0-0-0'),
	(10, 18, '0-0-0'),
	(10, 19, '0-0-0'),
	(10, 20, '0-0-0'),
	(10, 21, '0-0-0'),
	(11, 1, '467-94-0'),
	(11, 2, '468-98-0'),
	(11, 3, '0-0-0'),
	(11, 4, '465-86-0'),
	(11, 5, '463-1-1'),
	(11, 6, '1020-1-1'),
	(11, 7, '0-0-0'),
	(11, 8, '0-0-0'),
	(11, 9, '0-0-0'),
	(11, 10, '0-0-0'),
	(11, 11, '0-0-0'),
	(11, 12, '0-0-0'),
	(11, 13, '0-0-0'),
	(11, 14, '0-0-0'),
	(11, 15, '0-0-0'),
	(11, 16, '0-0-0'),
	(11, 17, '0-0-0'),
	(11, 18, '0-0-0'),
	(11, 19, '0-0-0'),
	(11, 20, '0-0-0'),
	(11, 21, '0-0-0'),
	(12, 1, '467-100-0'),
	(12, 2, '468-100-0'),
	(12, 3, '461-50-0'),
	(12, 4, '0-0-0'),
	(12, 5, '463-1-1'),
	(12, 6, '1020-1-1'),
	(12, 7, '465-100-0'),
	(12, 8, '0-0-0'),
	(12, 9, '0-0-0'),
	(12, 10, '0-0-0'),
	(12, 11, '0-0-0'),
	(12, 12, '0-0-0'),
	(12, 13, '0-0-0'),
	(12, 14, '0-0-0'),
	(12, 15, '0-0-0'),
	(12, 16, '0-0-0'),
	(12, 17, '0-0-0'),
	(12, 18, '0-0-0'),
	(12, 19, '0-0-0'),
	(12, 20, '0-0-0'),
	(12, 21, '0-0-0'),
	(12, 22, '0-0-0'),
	(12, 23, '0-0-0'),
	(12, 24, '0-0-0'),
	(12, 25, '0-0-0'),
	(12, 26, '0-0-0'),
	(12, 27, '0-0-0'),
	(12, 28, '0-0-0');
/*!40000 ALTER TABLE `rel_personaje_inventario` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_personaje_skill
CREATE TABLE IF NOT EXISTS `rel_personaje_skill` (
  `id_personaje` int(255) unsigned NOT NULL,
  `skill` mediumint(255) unsigned NOT NULL,
  `valor` smallint(255) unsigned DEFAULT NULL,
  PRIMARY KEY (`id_personaje`,`skill`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_personaje_skill: ~136 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_personaje_skill` DISABLE KEYS */;
INSERT INTO `rel_personaje_skill` (`id_personaje`, `skill`, `valor`) VALUES
	(1, 1, 40),
	(1, 2, 0),
	(1, 3, 75),
	(1, 4, 7),
	(1, 5, 78),
	(1, 6, 0),
	(1, 7, 0),
	(1, 8, 0),
	(1, 9, 0),
	(1, 10, 0),
	(1, 11, 0),
	(1, 12, 0),
	(1, 13, 0),
	(1, 14, 0),
	(1, 15, 0),
	(1, 16, 0),
	(1, 17, 0),
	(2, 1, 100),
	(2, 2, 100),
	(2, 3, 100),
	(2, 4, 100),
	(2, 5, 100),
	(2, 6, 100),
	(2, 7, 100),
	(2, 8, 100),
	(2, 9, 100),
	(2, 10, 100),
	(2, 11, 100),
	(2, 12, 100),
	(2, 13, 100),
	(2, 14, 100),
	(2, 15, 100),
	(2, 16, 100),
	(2, 17, 100),
	(3, 1, 100),
	(3, 2, 0),
	(3, 3, 0),
	(3, 4, 10),
	(3, 5, 2),
	(3, 6, 0),
	(3, 7, 0),
	(3, 8, 0),
	(3, 9, 0),
	(3, 10, 20),
	(3, 11, 3),
	(3, 12, 0),
	(3, 13, 100),
	(3, 14, 0),
	(3, 15, 69),
	(3, 16, 3),
	(3, 17, 100),
	(6, 1, 6),
	(6, 2, 0),
	(6, 3, 3),
	(6, 4, 2),
	(6, 5, 0),
	(6, 6, 0),
	(6, 7, 0),
	(6, 8, 0),
	(6, 9, 0),
	(6, 10, 0),
	(6, 11, 1),
	(6, 12, 0),
	(6, 13, 0),
	(6, 14, 0),
	(6, 15, 0),
	(6, 16, 0),
	(6, 17, 0),
	(9, 1, 0),
	(9, 2, 0),
	(9, 3, 1),
	(9, 4, 0),
	(9, 5, 0),
	(9, 6, 0),
	(9, 7, 0),
	(9, 8, 0),
	(9, 9, 0),
	(9, 10, 0),
	(9, 11, 0),
	(9, 12, 0),
	(9, 13, 0),
	(9, 14, 0),
	(9, 15, 0),
	(9, 16, 0),
	(9, 17, 0),
	(10, 1, 0),
	(10, 2, 0),
	(10, 3, 0),
	(10, 4, 0),
	(10, 5, 0),
	(10, 6, 0),
	(10, 7, 0),
	(10, 8, 0),
	(10, 9, 0),
	(10, 10, 0),
	(10, 11, 0),
	(10, 12, 0),
	(10, 13, 0),
	(10, 14, 0),
	(10, 15, 0),
	(10, 16, 0),
	(10, 17, 0),
	(11, 1, 2),
	(11, 2, 0),
	(11, 3, 2),
	(11, 4, 2),
	(11, 5, 2),
	(11, 6, 0),
	(11, 7, 0),
	(11, 8, 0),
	(11, 9, 0),
	(11, 10, 0),
	(11, 11, 0),
	(11, 12, 0),
	(11, 13, 0),
	(11, 14, 0),
	(11, 15, 0),
	(11, 16, 0),
	(11, 17, 0),
	(12, 1, 8),
	(12, 2, 0),
	(12, 3, 10),
	(12, 4, 1),
	(12, 5, 5),
	(12, 6, 0),
	(12, 7, 0),
	(12, 8, 0),
	(12, 9, 0),
	(12, 10, 0),
	(12, 11, 0),
	(12, 12, 0),
	(12, 13, 0),
	(12, 14, 0),
	(12, 15, 0),
	(12, 16, 0),
	(12, 17, 0);
/*!40000 ALTER TABLE `rel_personaje_skill` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_quest_npc
CREATE TABLE IF NOT EXISTS `rel_quest_npc` (
  `id_npc_start` smallint(6) DEFAULT NULL,
  `id_npc_end` smallint(6) DEFAULT NULL,
  `id_quest` smallint(6) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_quest_npc: ~2 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_quest_npc` DISABLE KEYS */;
INSERT INTO `rel_quest_npc` (`id_npc_start`, `id_npc_end`, `id_quest`) VALUES
	(645, 646, 1),
	(646, 645, 2);
/*!40000 ALTER TABLE `rel_quest_npc` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_quest_objetivo
CREATE TABLE IF NOT EXISTS `rel_quest_objetivo` (
  `id_quest` smallint(5) unsigned DEFAULT NULL,
  `tipo_objetivo` tinyint(3) unsigned DEFAULT NULL,
  `cantidad_objetivo` smallint(5) unsigned DEFAULT NULL,
  `objetivo` smallint(5) unsigned DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_quest_objetivo: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_quest_objetivo` DISABLE KEYS */;
INSERT INTO `rel_quest_objetivo` (`id_quest`, `tipo_objetivo`, `cantidad_objetivo`, `objetivo`) VALUES
	(1, 1, 20, 501);
/*!40000 ALTER TABLE `rel_quest_objetivo` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_quest_personaje_estado
CREATE TABLE IF NOT EXISTS `rel_quest_personaje_estado` (
  `id_quest` smallint(6) DEFAULT NULL,
  `id_personaje` tinyint(4) DEFAULT NULL,
  `estado` tinyint(4) DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_quest_personaje_estado: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_quest_personaje_estado` DISABLE KEYS */;
INSERT INTO `rel_quest_personaje_estado` (`id_quest`, `id_personaje`, `estado`) VALUES
	(1, 2, 5);
/*!40000 ALTER TABLE `rel_quest_personaje_estado` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_zona_efecto
CREATE TABLE IF NOT EXISTS `rel_zona_efecto` (
  `id_zona` mediumint(8) unsigned DEFAULT NULL,
  `id_efecto` mediumint(8) unsigned DEFAULT NULL,
  `evento` tinyint(3) unsigned DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_zona_efecto: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_zona_efecto` DISABLE KEYS */;
/*!40000 ALTER TABLE `rel_zona_efecto` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.rel_zona_npc
CREATE TABLE IF NOT EXISTS `rel_zona_npc` (
  `id_zona` mediumint(8) unsigned DEFAULT NULL,
  `id_npc` smallint(6) unsigned DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.rel_zona_npc: ~18 rows (aproximadamente)
/*!40000 ALTER TABLE `rel_zona_npc` DISABLE KEYS */;
INSERT INTO `rel_zona_npc` (`id_zona`, `id_npc`) VALUES
	(1, 501),
	(1, 501),
	(9, 596),
	(9, 596),
	(9, 596),
	(9, 596),
	(5, 507),
	(5, 502),
	(5, 502),
	(6, 502),
	(6, 502),
	(6, 507),
	(7, 502),
	(7, 502),
	(7, 507),
	(8, 502),
	(8, 502),
	(8, 507);
/*!40000 ALTER TABLE `rel_zona_npc` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.subasta
CREATE TABLE IF NOT EXISTS `subasta` (
  `objeto_id` int(11) DEFAULT NULL,
  `personaje_id` int(11) DEFAULT NULL,
  `cantidad` int(11) DEFAULT NULL,
  `buyout` int(11) DEFAULT NULL,
  `fecha_creacion` datetime DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.subasta: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `subasta` DISABLE KEYS */;
INSERT INTO `subasta` (`objeto_id`, `personaje_id`, `cantidad`, `buyout`, `fecha_creacion`) VALUES
	(386, 3, 1, 1, '2018-11-20 17:46:35');
/*!40000 ALTER TABLE `subasta` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.type_char
CREATE TABLE IF NOT EXISTS `type_char` (
  `id_npc` smallint(5) unsigned NOT NULL AUTO_INCREMENT,
  `charindex` smallint(5) unsigned NOT NULL,
  `head` smallint(5) unsigned NOT NULL,
  `body` smallint(5) unsigned NOT NULL,
  `animataque` smallint(5) unsigned NOT NULL,
  `weaponanim` smallint(5) unsigned NOT NULL,
  `cascoanim` smallint(5) unsigned NOT NULL,
  `shieldanim` smallint(5) unsigned NOT NULL,
  `fx` smallint(5) unsigned NOT NULL,
  `loops` smallint(5) unsigned NOT NULL,
  `heading` tinyint(3) unsigned NOT NULL DEFAULT 1,
  PRIMARY KEY (`id_npc`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.type_char: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `type_char` DISABLE KEYS */;
/*!40000 ALTER TABLE `type_char` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.type_npc
CREATE TABLE IF NOT EXISTS `type_npc` (
  `id` tinyint(4) NOT NULL AUTO_INCREMENT,
  `name` char(50) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=19 DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.type_npc: ~18 rows (aproximadamente)
/*!40000 ALTER TABLE `type_npc` DISABLE KEYS */;
INSERT INTO `type_npc` (`id`, `name`) VALUES
	(1, 'Comun'),
	(2, 'Revividor'),
	(3, 'GuardiaReal'),
	(4, 'Entrenador'),
	(5, 'Banquero'),
	(6, 'Noble'),
	(7, 'DRAGON'),
	(8, 'Timbero'),
	(9, 'Guardiascaos'),
	(10, 'ResucitadorNewbie'),
	(11, 'Pirata'),
	(12, 'MonsterDrag'),
	(13, 'Cirujano'),
	(14, 'Guardiafalso'),
	(15, 'Quest'),
	(16, 'Arbol'),
	(17, 'Yacimiento'),
	(18, 'Subastador');
/*!40000 ALTER TABLE `type_npc` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.type_npcai
CREATE TABLE IF NOT EXISTS `type_npcai` (
  `id` tinyint(4) NOT NULL AUTO_INCREMENT,
  `name` char(50) NOT NULL DEFAULT '0',
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=10 DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.type_npcai: ~9 rows (aproximadamente)
/*!40000 ALTER TABLE `type_npcai` DISABLE KEYS */;
INSERT INTO `type_npcai` (`id`, `name`) VALUES
	(1, 'Estatico'),
	(2, 'MueveAlAzar'),
	(3, 'NpcMaloAtacaUsersBuenos'),
	(4, 'NPCDEFENSA'),
	(5, 'GuardiasAtacanCriminales'),
	(6, 'NpcObjeto'),
	(7, 'SigueAmo'),
	(8, 'NpcAtacaNpc'),
	(9, 'NpcPathfinding');
/*!40000 ALTER TABLE `type_npcai` ENABLE KEYS */;

-- Volcando estructura para tabla aodrag.zona
CREATE TABLE IF NOT EXISTS `zona` (
  `id` mediumint(8) unsigned NOT NULL AUTO_INCREMENT,
  `nombre` char(50) DEFAULT NULL,
  `mapa` tinyint(3) unsigned DEFAULT NULL,
  `x1` tinyint(3) unsigned DEFAULT NULL,
  `x2` tinyint(3) unsigned DEFAULT NULL,
  `y1` tinyint(3) unsigned DEFAULT NULL,
  `y2` tinyint(3) unsigned DEFAULT NULL,
  `permisos` smallint(5) unsigned DEFAULT 0,
  `grh` mediumint(9) DEFAULT 0,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=utf8;

-- Volcando datos para la tabla aodrag.zona: ~0 rows (aproximadamente)
/*!40000 ALTER TABLE `zona` DISABLE KEYS */;
INSERT INTO `zona` (`id`, `nombre`, `mapa`, `x1`, `x2`, `y1`, `y2`, `permisos`, `grh`) VALUES
	(1, 'asdf', 1, 1, 1, 1, 1, 0, 0);
/*!40000 ALTER TABLE `zona` ENABLE KEYS */;

/*!40101 SET SQL_MODE=IFNULL(@OLD_SQL_MODE, '') */;
/*!40014 SET FOREIGN_KEY_CHECKS=IF(@OLD_FOREIGN_KEY_CHECKS IS NULL, 1, @OLD_FOREIGN_KEY_CHECKS) */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
