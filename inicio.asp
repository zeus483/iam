 <!DOCTYPE html>
<html lang="es">
<head>
	<meta charset="utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE-edge">
	<meta name="viewport" content="width=device-width, initial-scale=1">
	<!-- Las etiquetas meta anteriores deben ir siempre en el head antes de cualquier otra etiqueta -->
	<title>Asset Management</title>
	<link rel="shortcut icon" href="./ScriptsJQuery/iconos/Iconos Bancolombia/LogoBancolombia.ico">
	<!-- Estilos -->
	<!-- Bootstrap -->
	<link rel="stylesheet" type="text/css" href="../ScriptsJQuery/bootstrap/css/bootstrap.min.css">
	<link rel="stylesheet" href="../ScriptsJQuery/bootstrap/css/bootstrap-glyphicons.css">
	<link rel="stylesheet" href="../ScriptsJQuery/bootstrap/css/bootstrap.fixheadertable.css">
	<!-- JQwidgets -->
	<link rel="stylesheet" type="text/css" href="../ScriptsJQuery/jqwidgets/jqwidgets/styles/jqx.base.css">
	<link rel="stylesheet" type="text/css" href="../ScriptsJQuery/jqwidgets/jqwidgets/styles/jqx.bootstrap.css">
	<link rel="stylesheet" type="text/css" href="../ScriptsJQuery/jqwidgets/jqwidgets/styles/jqx.metrodark.css">
	<!-- Librerias -->
	<!-- JQuery -->
	<script type="text/javascript" src="../ScriptsJQuery/bootstrap/js/jquery-2.1.0.min.js"></script>
	<script type="text/javascript" src="../ScriptsJQuery/jquery.easing.1.3.js"></script>
	<!-- Bootstrap -->
	<script type="text/javascript" src="../ScriptsJQuery/bootstrap/js/bootstrap.min.js"></script>
	<!-- JQwidgets -->
	<script type="text/javascript" src="../ScriptsJQuery/jqwidgets/jqwidgets/jqxcore.js"></script>
	<script type="text/javascript" src="../ScriptsJQuery/jqwidgets/jqwidgets/jqx-all.js"></script>
	<!-- Configuracion regional JQwidgets -->
	<script type="text/javascript" src="../ScriptsJQuery/jqwidgets/jqwidgets/globalization/globalize.js"></script>
	<script type="text/javascript" src="../ScriptsJQuery/jqwidgets/jqwidgets/globalization/globalize.culture.es-CO.js"></script>
	<!-- Alasql -->
	<script type="text/javascript" src="../ScriptsJQuery/alasql/alasql.min.js"></script>
	<script type="text/javascript" src="../ScriptsJQuery/alasql/xls.core.min.js"></script>
	<script type="text/javascript" src="../ScriptsJQuery/alasql/xlsx.core.min.js"></script>
	<!-- festivosJson -->
	<script type="text/javascript" src="../ScriptsJQuery/festivosJson/festivos.js"></script>
	<!-- jsDate -->
	<script type="text/javascript" src="../ScriptsJQuery/jsDate.js"></script>
	<!-- shortcut js -->
	<script type="text/javascript" src="../ScriptsJQuery/Shorcuts/Shorcuts.js" charset="utf-8"></script>
	<!-- numero a letras js -->
	<script type="text/javascript" src="../ScriptsJQuery/numeroALetras.js" charset="utf-8"></script>
	<!-- numeralFormat -->
	<script type="text/javascript" src="../ScriptsJQuery/numeralFormat/numeral.min.js" charset="utf-8"></script>
	<style type="text/css">
		* {
			box-sizing: border-box;
		}
		#BarraColapsable, .jqx-menu-ul-metrodark, .jqx-menu-dropdown-metrodark {
			/* background-color: #00448c !important; */
			background-color: #000 !important;
		}
		.jqx-menu-dropdown {
			border-radius: 5px;
		}
		li[tipoMenu*="principal"] {
			height: 34px;
			font-size: 16px !important;
			padding-top: 14px;
			padding-bottom: 0px;
		}
		#BarraColapsable > ul > li:hover {
			border-radius: 5px;
			background-color: #337AB7;
		}
		#BarraColapsable > ul > li {
			font-size: 12px;
			/* background-color: #00448c !important; */
			background-color: #000 !important;
			color: #FFF !important;
			margin: 0px !important;
		}
		#BarraColapsable > ul > a:hover {
			color: #FFF !important;
			font-weight: bold;
			border-radius: 5px;
			background-color: #337AB7 !important;
		}
		#BarraColapsable > ul > a {
			font-size: 12px;
			color: #FFF !important;
			background-color: #00448c !important;
		}
		.navbar{
			min-height: 40px;
			height: 40px;
		}
		.navbar-nav > li {
			padding: 0px;
			height: 38px;
		}
		.navbar-nav > li > a {
			padding: 10px 10px;
			height: 38px;
			vertical-align: middle;
			color: #FFF !important;
			font-size: 13px;
		}
		.navbar-brand {
		  height: 39px;
		  padding: 10px 10px;
		  font-size: 16px;
		  line-height: 18px;
		}
		.center {
			display: block;
			margin: auto;
		}
		.margin-right-15 {
			margin-right: 15px !important;
		}
		.fleft-mr10-pd4 {
			float: left;
			margin-right: 10px !important;
			padding-top: 4px;
			margin-top: 5px !important;
		}
		.fleft {
			float: left;
			margin-top: 5px !important;
		}
		.lbflet {
			float: left;
			margin-right: 5px;
			padding-top: 4px;
			margin-top: 5px !important;
			height: 25px;
		}
		.table-fsize-9 {
			font-size: 8pt;
		}
		.table-fsize-9 th {
			text-align: center;
		}
		.table-fsize-9 td {
			text-align: right;
		}
		.table-fsize-9 .text-center {
			text-align: center;
		}
		.tb-mg-bt-0 {
			margin-bottom: 0px !important;
		}
		.tb-danger {
			background-color: red;
			color: white;
		}
		table.table-fixed>tbody {
		  max-height: 115px;
		}
		.grid-cell-estado-liberado {
			background: linear-gradient(#00C853, #00E676);
			color: black;
		}
		.grid-cell-estado-revisado {
			background: linear-gradient(#FFEA00, #FFFF00);
			color: black;
		}
		.grid-cell-estado-sinSaldo {
			background: linear-gradient(#EC407A, #E91E63);
			color: white;
		}
		.grid-cell-estado-Finalizado {
			background: linear-gradient(#304FFE, #304FFE);
			color: white;
		}
		.grid-cell-estado-rechazo {
			background: linear-gradient(red, red);
			color: white;
		}
		.grid-cell-estado-pendientePago {
			background: #FFF9C4;
			color: black;
		}
		.italic-bold-font {
			font-style: italic;
			/* font-weight:bold; */
		}
		.No_Seleccionar_Texto {
			-webkit-user-select: none;
			-moz-user-select: none;
			-khtml-user-select: none;
			-ms-user-select:none;
		} 
	</style>
</head>
<body>
	<div id="TopPage"></div>
	<%
		dim Usuario
		Usuario = split(Request.ServerVariables("AUTH_USER"),"\")
  'Pipe: cuando se salga a producción la siguiente
  'línea debe quedar sin comentario.		
		'UsuarioActual = ucase(Usuario(1))
  'Pipe: cuando se salga a producción la siguiente
  'línea se debe eliminar.
		UsuarioActual = "juaosori"
	%>

	<div style="display: flex; justify-content: center; width: 100%;">
	<div id="BarraColapsable" style="visibility: hidden; background-color: #00448c; color: #FFF; position: sticky; top: 0px; padding: 0px; z-index: 999999;">
		<ul style="padding: 0px;">
			<li style="background-color: #FFF !important; color: #000 !important; margin: 0px !important; height: 48px;">
				<a style="background-color: #FFF; font-size: 16px;" href="javascript:void(0)" class="navbar-brand"
    img src = "./ScriptsJQuery/iconos/Iconos Bancolombia/LogoBancolombia.ico"  height=32 style="float: left; position: relative; top: -3px; margin-left: 6px; border-left: 2px solid #EFF0F1;"></a>
   <a style="color: #000; font-size: 14px;">Asset Management</a>
    <img src = "./ScriptsJQuery/iconos/Iconos Bancolombia/LogoBancolombia.ico" height=32 style="float: left; position: relative; top: -3px; margin-left: 6px; border-left: 2px solid #EFF0F1;"></a> 
			
   </li>
			<li tipoMenu="principal" tipoPerfil="admins trader" style="display: none;">Trader
				<ul style="width: 180px;">
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Ingreso TRM</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Admon portafolios</a></li>
					<!-- <li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Cargue insumos</a></li> -->
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Cuadre individual</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Cuadre masivo</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Estado de operaciones</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Controles</a></li>
				 <!-- <li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Admon matriz movimientos</a></li> -->
				</ul>
			</li>
			<li tipoMenu="principal" tipoPerfil="admins back-end" style="display: none;">Tesoreria MK
				<ul>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Cargue insumos</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Gestión repos depósitos remunerados</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Gestión operaciones</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Cartas repo / depósitos remunerados</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Traslados</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Estado de operaciones</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Admon matriz movimientos</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Operaciones manuales BE</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Eliminación / reversión operaciones</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Control Murex</a></li>
				</ul>
			</li>
			<li tipoMenu="principal" tipoPerfil="admins cumplimiento" style="display: none;">Cumplimiento
				<ul>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Gestión repos depósitos remunerados</a></li>
				</ul>
			</li>
			<li tipoMenu="principal" tipoPerfil="admins" style="display: none;">Administración
				<ul>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Admon usuarios</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Admon matriz movimientos</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Admon bancos</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Portafolios no gestionados</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Eliminación / reversión operaciones</a></li>
				</ul>
			</li>
			<li tipoMenu="principal" tipoPerfil="admins lectura back-end" style="display: none;">Informes
				<ul>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Estado de operaciones</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Detalles repo / depósitos remunerados</a></li>
     <li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Contabilidad Valores</a></li>
     <li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Otros bancos</a></li>
				</ul>
			</li>
			<li tipoMenu="principal" tipoPerfil="admins" style="desplay: none;">intencionesam
				<ul>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Cargue Insumos Intenciones</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Mis Ordenes</a></li>
					<li tipoMenu="secundario"><a href="javascript:void(0)" class="elmenu">Monitor</a></li>
				</ul>
			</li>
			
		</ul>
	</div>
	</div>

	<div class="navbar-fixed-bottom alert alert-danger alert-dismissible" id="AlertasDiv" style="display: none; max-height: 400px; height: auto; overflow-y: auto; z-index: 999999;"  role="alert">
		<button type="button" class="close" data-hide="alert" aria-label="Cerrar"><span aria-hidden="true">&times;</span></button>
		<div id="contenidoAlerta"></div>
	</div>

	<div class="container-fluid" style="margin-top: 10px; visibility: hidden;" id="ContenedorDeCarga">
		<div id="contenidoContCarga">
  <img  class="center img-thumbnail img-responsive" src="./ScriptsJQuery/iconos/Iconos Bancolombia/logo_inicio_banco.png" style="margin-top: 20px; height: 490px; width: 1220px; !important;"> 
		</div>
	</div>
	<div id= "pieLineaEtica" style="text-align:center;padding-top:2px;font-size: 10px; font-family: Verdana;">
		La información que aparece en este sitio es exclusiva para los empleados del Grupo Bancolombia. Por consiguiente, no va dirigida ni tiene como destinatario a terceras personas o entidades. <br> Ningún colaborador está autorizado para ponerla en circulación a través de cualquier medio, ni divulgarla a personas distintas al grupo.<br />
		<a href="http://vsc.bancolombia.corp/vsa/Seguridad/lineaEtica/Paginas/default.aspx" target='_blank' ><img style="margin-left:10px" alt="" border="0" src="../ScriptsJQuery/iconos/imagenes/linea_etica.jpg"/></a>
	</div>

	<input type="hidden" id="UsuarioIngresa" value="<%=UsuarioActual%>">
	<input type="hidden" id="NombreUsuarioIngresa">
	<input type="hidden" id="PerfilUsuarioIngresa">
	<div id="jqxLoader"></div>
	<script type="text/javascript">
		$(document).ready(function(){
			//$('#UsuarioIngresa').val(window.prompt("Ingrese su usuario de red"));
			//Validación de usuario
			
			$.ajax({
				url: 'Consultas.asp',
				type: 'post',
				dataType: 'json',
				data: {
					opcion: 'ValidaUsuario',
					Usuario: $('#UsuarioIngresa').val()
				},
				success: function(data){
					function capitalizarClaves(obj) {
					const nuevo = {};
					for (const key in obj) {
						if (obj.hasOwnProperty(key)) {
							const nuevaClave = key.charAt(0).toUpperCase() + key.slice(1);
							nuevo[nuevaClave] = obj[key];
						}
					}
					return nuevo;
					}
					if (data.length > 0) {
						data[0] = capitalizarClaves(data[0]);
						$('#NombreUsuarioIngresa').val(data[0].Nombre);
						$('#PerfilUsuarioIngresa').val(data[0].Perfil);
						$('#BarraColapsable').jqxMenu({ width: '98%', height: '50px', popupZIndex: 999999, theme: 'metrodark', autoOpen: false });

						$('#BarraColapsable').css({
							opacity: 0,
							height: 0,
							width: 0,
							left: '50%',
							right: 5,
							'background-color': '#00448C',
							'border-radius': '5px',
							visibility: 'visible'
						}).animate({
							opacity: 1,
							height: '50px',
							width: '98%',
							left: 0,
						});

						var w = $(document).width() * 0.98;
						$('#contenidoContCarga').css({
							width: 0,
							height: 0,
							opacity: 0,
							visibility: 'visible',
						}).animate({
							opacity: 1,
							width: w,
							height: 'auto'
						}, {
							duration: 1000,
							specialEasing: {
								width: 'easeInQuart',
								height: 'easeInQuart'
							}
						});

						switch(data[0].Perfil){
							case 1:
								$('[tipoPerfil*="trader"]').css('display', 'block');
								break;
							case 2:
								$('[tipoPerfil*="back-end"]').css('display', 'block');
								break;
							case 4:
								$('[tipoPerfil*="admins"]').css('display', 'block');
								break;
							case 5:
								$('[tipoPerfil*="lectura"]').css('display', 'block');
								break;
							case 6:
								$('[tipoPerfil*="cumplimiento"]').css('display', 'block');
								break;
						}
						switch(data[0].Perfil){
							case 1:
							case 2:
							case 4:
							case 5:
							case 6:
							$(function() {
								$(window).scroll(function() {
									$("#BarraColapsable").css('opacity', 1 - $(window).scrollTop() / 50);
									var opa = $("#BarraColapsable").css('opacity');
									if (opa == 0) {
										$('#BarraColapsable').css('visibility', 'hidden');
									} else {
										$('#BarraColapsable').css('visibility', 'visible');
									}
								});
							});
							break;
						};
					} else {
						$('#contenidoContCarga').html("<center><font size='4' face='calibri'><b>El usuario " + $('#UsuarioIngresa').val() + " no tiene permisos en esta herramienta. Comuniquese con su jefe para la debida parametrización.</b></font></center>");
						$('#contenidoContCarga').css('visibility', 'visible')
					}
				},
				error: function(){
					
					$('#contenidoAlerta').html('<span class="glyphicon glyphicon-warning-sign"></span>&nbsp;&nbsp;<b>Importante</b><br>No hay conexión con el servidor');
					if (!$('#AlertasDiv').hasClass('alert-danger')) {
						if ($('#AlertasDiv').hasClass('alert-success')) {
							$('#AlertasDiv').removeClass('alert-success');
						}
						$('#AlertasDiv').addClass('alert-danger');
					}
					$('#AlertasDiv').css('display', 'block');
					setTimeout(escondeAlerta, 4000);
				}
			});
		});

		var lafec = new Date();
		var y = lafec.getFullYear();
		var festivos = getColombiaHolidaysByYear(y);
		// console.log(festivos);

		function encuentraDiaHabil()
  {
     let lafechaActual = new Date();
     lafechaActual.setDate(lafechaActual.getDate() + 1);
     // let lafechaActual = '2021-01-01';
     let diaSemana = lafechaActual.getDay();

     if (diaSemana == 6) { // sabado
      lafechaActual.setDate(lafechaActual.getDate() + 2);
     }
     if (diaSemana == 0) { // domingo
      lafechaActual.setDate(lafechaActual.getDate() + 1);
     }

     let dia = (lafechaActual.getDate()).toString();
     let mes = (lafechaActual.getMonth() + 1).toString();
     let yy = lafechaActual.getFullYear().toString();
     if (dia.length == 1) {
      dia = '0' + dia;
     }
     if (mes.length == 1) {
      mes = '0' + mes;
     }
     let lafechaActualForm = yy + '-' + mes + '-' + dia;
     // console.log(lafechaActualForm)
     //se valida si el día resultante es festivo
     for (let index = 0; index < festivos.length; index++) {
      if (festivos[index].festivo == lafechaActualForm) {
       lafechaActual.setDate(lafechaActual.getDate() + 1);
       let diaSemana = lafechaActual.getDay();

       if (diaSemana == 6) { // sabado
        lafechaActual.setDate(lafechaActual.getDate() + 2);
       }
       if (diaSemana == 0) { // domingo
        lafechaActual.setDate(lafechaActual.getDate() + 1);
       }
      }
     }
     dia = (lafechaActual.getDate()).toString();
     mes = (lafechaActual.getMonth() + 1).toString();
     yy = lafechaActual.getFullYear().toString();
     if (dia.length == 1) {
      dia = '0' + dia;
     }
     if (mes.length == 1) {
      mes = '0' + mes;
     }
     return lafechaActualForm = yy + mes + dia;
     // console.log(lafechaActualForm)
		}
		
		function normalize(texto) {
		 	var compara = "ÃÀÁÄÂÈÉËÊÌÍÏÎÒÓÖÔÙÚÜÛãàáäâèéëêìíïîòóöôùúüûÑñÇç",
	      	cambia   = "AAAAAEEEEIIIIOOOOUUUUaaaaaeeeeiiiioooouuuuNncc",
	      	normalizado = "", index;

	 		for (var i = 0; i < texto.length; i++) {
	 			index = compara.indexOf(texto.charAt(i));
	 			if (index >=0) {
	 				normalizado += cambia[index];
	 			} else {
	 				normalizado += texto.charAt(i);
	 			}
	 		}

		  	normalizado = normalizado.replace(/[^A-Za-z0-9]+/g, '' );
		  	return normalizado
		};

		function normalize2(texto) 
  {
		 	var compara = "ÃÀÁÄÂÈÉËÊÌÍÏÎÒÓÖÔÙÚÜÛãàáäâèéëêìíïîòóöôùúüûÑñÇç",
	      	cambia   = "AAAAAEEEEIIIIOOOOUUUUaaaaaeeeeiiiioooouuuuNncc",
	      	normalizado = "", index;

	 		for (var i = 0; i < texto.length; i++) 
    {
       index = compara.indexOf(texto.charAt(i));
       if (index >=0) 
       {
         normalizado += cambia[index];
       } 
       else 
       {
         normalizado += texto.charAt(i);
       }
	 		}
     /*Remueve caracteres diferentes de:
       i) letras (en mayúsculas o minúsculas)
       ii) números de 0 a 9
       iii) puntos (.)
       iv) tíldes
       v) símbolo #
       vi) espacios */
		  	normalizado = normalizado.replace(/[^A-Za-z0-9.\´\# ]+/g, '' );
		  	return normalizado
		};

		function cierraVentanasFlotantes(){
			$('#contenidoAlerta').empty();
			$('#AlertasDiv').css('display', 'none');
		};

		$('.navbar-brand').click(function(event){
			cierraVentanasFlotantes();
			$('#contenidoContCarga').html('<img class="center img-thumbnail img-responsive" src="./ScriptsJQuery/iconos/Iconos Bancolombia/logo_inicio_banco.png" style="margin-top: 20px; height: 490px; width: 1220px;">');
			var w = $(document).width() * 0.98;
			$('#contenidoContCarga').css({
				width: 0,
				height: 0,
				opacity: 0,
				visibility: 'visible',
			}).animate({
				opacity: 1,
				width: w,
				height: 'auto'
			}, {
				duration: 1000,
				specialEasing: {
					width: 'easeInQuart',
					height: 'easeInQuart'
				}
			});
		});

		$("#jqxLoader").jqxLoader({ width: 100, height: 60, imagePosition: 'top', text: 'Cargando...' });
		$('#BarraColapsable').on('itemclick', function(event){
			var element = event.args;
			var eltipoMenu =  $('#' + element.id)[0].attributes.tipoMenu;
			if (eltipoMenu) {
				eltipoMenu = eltipoMenu.nodeValue;
				if (eltipoMenu.indexOf('secundario') >= 0) {
					try	{
						$('#ttVencfidu').jqxTooltip('destroy');
						$('#ttAcciones').jqxTooltip('destroy');
						$('#ttCVFiduciaria').jqxTooltip('destroy');
						$('#ttVencValores').jqxTooltip('destroy');
						$('#ttMesa').jqxTooltip('destroy');
						$('#ttComisiones').jqxTooltip('destroy');
						$('#ttWebFVal').jqxTooltip('destroy');
						$('#ttOpVigRF').jqxTooltip('destroy');
						$('#ttFuturos').jqxTooltip('destroy');
						$('#ttSwap').jqxTooltip('destroy');
						$('#ttForward').jqxTooltip('destroy');
						$('#ttPosPorfin').jqxTooltip('destroy');
						$('#ttGantFidu').jqxTooltip('destroy');
						$('#ttPosGantFidu').jqxTooltip('destroy');
						$('#ttGantVal').jqxTooltip('destroy');
						$('#ttPosGantVal').jqxTooltip('destroy');
						$('#ttPosOYD').jqxTooltip('destroy');
						$('#ttVencMurex').jqxTooltip('destroy');
						$('#ttVencOyD').jqxTooltip('destroy');
						$('#ttCumplOyD').jqxTooltip('destroy');
						$('#ttOperOyD').jqxTooltip('destroy');
						$('#ttCtasClear').jqxTooltip('destroy');
						$('#ttSaldosFidu').jqxTooltip('destroy');
						$('#ttSaldosCarterasColectivas').jqxTooltip('destroy');
					} catch(err){
					}
					cierraVentanasFlotantes();
					var primerMensaje = 0;
					$('#jqxLoader').jqxLoader('open');
     /*Pipe: en esta parte captura el texto
       del elemento donde se haya hecho clic.*/
					var laPagina = normalize(element.innerText);
					var elperfil = $('#PerfilUsuarioIngresa').val();
					// excepciones para paginas en construccion, solo muestra la pagina real a los administradores, los demás verán la pagina con el mensaje de en construccion
					// if (laPagina == 'OperacionesManualesBE' && elperfil != 4) {
					// 	laPagina = 'EnConstruccion';
					// }
					$('[libera-eventos="true"]').off();
					$('[libera-eventos="true"]').remove();
					shortcut.removeAll();
					$('.jqx-window-modal').remove();
					$('#contenidoContCarga').empty();
					$('#contenidoContCarga').remove();
					$('#ContenedorDeCarga').html('<div id="contenidoContCarga"></div>');
					$(document).scrollTop( $("#TopPage").offset().top );
					if (laPagina === '') {
						$('#contenidoContCarga').css('display', 'none');
						$("#jqxLoader").jqxLoader('close');
						return false
					};

					var d = new Date();
					var h = d.getHours().toString();
					var m = d.getMinutes().toString();
					var s = d.getSeconds().toString();
					var r = h + m + s;
					var w = $(document).width();

					if ($('#BarraColapsable').hasClass('in')) {
						$('#BarraColapsable').removeClass('in')
					}

					$('#contenidoContCarga').load(laPagina + '.html?r=' + r, function(response, status, xhr) {
						if (status === 'error' && primerMensaje === 0) {
							$('#contenidoAlerta').html('<span style="font-size: 16px;" class="glyphicon glyphicon-warning-sign"></span>&nbsp;&nbsp;<b>Importante</b><br>Ha ocurrido un error con la carga de la página. Por favor inténtelo de nuevo.');
							if (!$('#AlertasDiv').hasClass('alert-danger')) {
								if ($('#AlertasDiv').hasClass('alert-success')) {
									$('#AlertasDiv').removeClass('alert-success');
								}
								$('#AlertasDiv').addClass('alert-danger');
							}
							$('#AlertasDiv').css('display', 'block');
							setTimeout(escondeAlerta, 4000);
							primerMensaje += 1
							$('#contenidoContCarga').html('<img class="center img-thumbnail img-responsive" src="./ScriptsJQuery/iconos/Iconos Bancolombia/logo_inicio_banco.png" style="margin-top: 20px; height: 490px; width: 1220px;">');
							var w = $(document).width() * 0.98;
							$('#contenidoContCarga').css({
								width: 0,
								height: 0,
								opacity: 0,
								visibility: 'visible',
							}).animate({
								opacity: 1,
								width: w,
								height: 'auto'
							}, {
								duration: 1000,
								specialEasing: {
									width: 'easeInQuart',
									height: 'easeInQuart'
								}
							});
						} else if (status !== 'error') {
							$('#contenidoContCarga').css({
								height: 0,
								opacity: 0,
								visibility: 'visible',
							}).animate({
								opacity: 1,
								height: 'auto'
							}, {
								duration: 2000,
								specialEasing: {
									height: 'easeInBounce'
								}
							});
						}
						$("#jqxLoader").jqxLoader('close');
					});
				}
			}
		});
/*		$('.elmenu, [tipoMenu="secundario"]').click(function(event){
		});
*/
		function escondeAlerta(){
			$('#AlertasDiv').css('display', 'none');
		};

		$(function(){
		    $("[data-hide]").on("click", function(){
		        $("." + $(this).attr("data-hide")).hide();
		        // -or-, see below
		        // $(this).closest("." + $(this).attr("data-hide")).hide();
		    });
		});

		function escondeMA(){
			$("span:contains('www.jqwidgets.com')").html('');
		};

		var getLocalization = function () 
  {
   var localizationobj = {};
   localizationobj.pagergotopagestring = "Ir a:";
   localizationobj.pagershowrowsstring = "Mostrar Filas:";
   localizationobj.pagerrangestring = " de ";
   localizationobj.pagernextbuttonstring = "Siguiente";
   localizationobj.pagerpreviousbuttonstring = "Anterior";
   localizationobj.pagerlastbuttonstring = "Ultima";
   localizationobj.pagerfirstbuttonstring = "Primera";
   localizationobj.sortascendingstring = "Ordenar ascendemte";
   localizationobj.sortdescendingstring = "Ordenar descendente";
   localizationobj.sortremovestring = "Remover ordenado";
   localizationobj.percentsymbol = "%";
   localizationobj.currencysymbol = "$";
   localizationobj.currencysymbolposition = "before";
   localizationobj.decimalseparator = ".";
   localizationobj.thousandsseparator = ",";
   localizationobj.filterstring = "Filtrar";
   localizationobj.filterclearstring = "Limpiar filtro";
			localizationobj.filtershowrowstring = "Mostrar filas donde:";
			localizationobj.filtershowrowdatestring = "Mostrar filas donde fecha:";
			localizationobj.filterorconditionstring = "O";
			localizationobj.filterandconditionstring = "Y";
			localizationobj.filterselectallstring = "(Seleccionar Todos)";
			localizationobj.filterchoosestring = "Seleccione:";
			localizationobj.filterstringcomparisonoperators = ['vacio', 'no vacio', 'contiene', 'contiene(coincide mayusculas)',
				'no contiente', 'no contiente(coincide mayusculas)', 'Comienza con', 'Comienza con(coincide mayusculas)',
				'Termina con', 'Termina con(coincide mayusculas)', 'igual', 'igual(coincide mayusculas)', 'nulo', 'no nulo'];
			localizationobj.filternumericcomparisonoperators = ['igual', 'no es igual', 'menor que', 'menor o igual que', 'mayor que', 'mayor o igual que', 'nulo', 'no nulo'];
			localizationobj.filterdatecomparisonoperators = ['igual', 'no es igual', 'menor que', 'menor o igual que', 'mayor que', 'mayor o igual que', 'nulo', 'no nulo'];
			localizationobj.filterbooleancomparisonoperators = ['igual', 'no es igual'];
			localizationobj.validationstring = "El valor ingresado no es valido";
			localizationobj.emptydatastring = "No hay datos para mostrar";
			localizationobj.filterselectstring = "Seleccione Filtro";
			localizationobj.loadtext = "Carga...";
			localizationobj.clearstring = "Limpiar";
			localizationobj.todaystring = "Hoy";
			localizationobj.groupsheaderstring = "Arrastre una columna para agrupar";
			localizationobj.groupbystring = "Agrupar por esta columna";
			localizationobj.groupremovestring = "Quitar de grupos";
   var days = 
   {
       // full day names
       names: ["Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado"],
       // abbreviated day names
       namesAbbr: ["Dom", "Lun", "Mar", "Mier", "Jue", "Vie", "Sab"],
       // shortest day names
       namesShort: ["Do", "Lu", "Ma", "Mi", "Ju", "Vi", "Sa"]
   };
   localizationobj.days = days;
   var months = {
       // full month names (13 months for lunar calendards -- 13th month should be "" if not lunar)
       names: ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre", ""],
       // abbreviated month names
       namesAbbr: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic", ""]
   };
   localizationobj.months = months;
      var patterns = {
          d: "dd/MM/yyyy",
          D: "dddd, d. MMMM yyyy",
          t: "HH:mm",
          T: "HH:mm:ss",
          f: "dddd, d. MMMM yyyy HH:mm",
          F: "dddd, d. MMMM yyyy HH:mm:ss",
          M: "dd MMMM",
          Y: "MMMM yyyy"
      }
      localizationobj.patterns = patterns;
      return localizationobj;
		};

		//funcion para convertir las fechas cuando excel las devuelve como código de fecha; ejemplo "2009-10-18 = 40105", al traer 40105, se debe transformar de nuevo la fecha en 2009-10-18
		alasql.fn.datetime = function(fecha) {
			if (fecha === '' || !fecha) {return '';};

			var fec = new Date((fecha - (25567 + 1))*86400*1000);
			var d = fec.getDate();
			var m = fec.getMonth() + 1;
			var y = fec.getFullYear();

			if (d.toString().length === 1) {
				d = '0' + d;
			}
			if (m.toString().length === 1) {
				m = '0' + m;
			}

			if ($.isNumeric(fecha)) {
				return y + '-' + m + '-' + d //(d + '/' + m + '/' + y).toString();
			} else {
				return $.trim(fecha.toString());
			}

		};

		//funcion para validar campos nulos en alasql
		alasql.fn.CheckNullUndefinedV =  function(val, tip){
			//nulo y numero
			if (val == null && tip == 1) {
				return 0;
			} else {
				return val;
			}
			//nulo y string
			if (val == null && tip == 2) {
				return "";
			} else {
				return val;
			}
		};
  
  /*Pipe: con esta función completa a 3 dígitos
    los portafolios de Porfin.*/

		//funcion para validar los digitos de los portafolios
		alasql.fn.checkPorta2Digits = function(val){
			if (val) {
				var port = $.trim(val.toString());
				port = $.trim(port.substring(0, 3).replace('-', ''));
				if (port.length < 3) {
					port = port + 'R';
				}
				return port;
			} else {
				return val;
			}
		};

		// funcion para tratar las fechas en formato 2018-09-10T05:00:00.000Z
		alasql.fn.fechaStd = function (val) {
			if (val === '' || !val) {return '';};

			var fec = new Date(val);
			var d = fec.getDate();
			var m = fec.getMonth() + 1;
			var y = fec.getFullYear();

			if (d.toString().length === 1) {
				d = '0' + d;
			}
			if (m.toString().length === 1) {
				m = '0' + m;
			}

				return y + '-' + m + '-' + d //(d + '/' + m + '/' + y).toString();
		 };

		alasql.fn.fechaStd2 = function (val) 
  {
     if (val === '' || !val) {return '';};
     // 2018.01.01
     var d = val.substring(8);
     var m = val.substring(5, 7);
     var y = val.substring(0, 4);

     if (d.toString().length === 1) {
      d = '0' + d;
     }
     if (m.toString().length === 1) {
      m = '0' + m;
     }

      return y + '-' + m + '-' + d //(d + '/' + m + '/' + y).toString();
		 };

		alasql.fn.numeral = numeral;
		
		var getSelected = function(){
		    var t = '';
		    if(window.getSelection) {
		        t = window.getSelection();
		    } else if(document.getSelection) {
		        t = document.getSelection();
		    } else if(document.selection) {
		        t = document.selection.createRange().text;
		    }
		    return t;
		};

		function pad(str, max) {
		    str = str.toString();
		    return str.length < max ? pad("0" + str, max) : str;
		};

		function addnsbp(str, max) {
		    str = str.toString();
		    return str.length < max ? addnsbp(" " + str, max) : str;
		};

		function addnsbpend(str, max) {
		    str = str.toString();
		    return str.length < max ? addnsbpend(str + " ", max) : str;
		};

		function generaId(port, origen){
			var fec = new Date();
			var d = pad(fec.getDate(), 2);
			var m = pad(fec.getMonth() + 1, 2);
			var y = fec.getFullYear();
			var h = pad(fec.getHours(), 2);
			var mm = pad(fec.getMinutes(), 2);
			var s = pad(fec.getSeconds(), 2);
			var ms = pad(fec.getMilliseconds(), 4);

			return y + m + d + '-' + port + '-' + h + mm + s + ms + '-' + origen;
		};

		function generaIdRVL(port, cont, hop, origen){
			var fec = new Date();
			var d = pad(fec.getDate(), 2);
			var m = pad(fec.getMonth() + 1, 2);
			var y = fec.getFullYear();
			var c = cont.substring(0, 10).replace(/[ \[\]]/g, "");

			return y + m + d + '-' + port + '-' + c + '-' + hop + '-' + origen;
		};
		var sourceFormatoNumeros = [];
		var sourceFormatoNumerosDA = new $.jqx.dataAdapter(sourceFormatoNumeros);
 
  /*Pone formato a número: 2 decimales, coma
    como separador de miles y rojo para valores
    negativos dentro de cada celda.*/
  function PonerFormatoNumero (valor)
  {
     //Crea condiciones para centrado vertical.
     var style = 'display: flex; align-items: center; justify-content: flex-end; height: 100%;';   
     //Valida si valor es negativo.
     if (valor < 0) 
        // return "<span style='display: block; text-align:right; color:red;'>$" + valor.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "</span>";
        return "<span style='" + style + "color:red;'>$" + valor.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "</span>";
     else
        // return "<span style='display: block; text-align:right;'>$" + valor.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "</span>";
        return "<span style='" + style + "'>$" + valor.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + "</span>";
  }
  /*Pone formato a número: 2 decimales, coma
    como separador de miles y rojo para valores
    negativos abajo del gridview en el total.*/
  function PonerFormatoValorTotal (valor)
  {
     //Crea condiciones para centrado vertical.
     var style = 'display: flex; align-items: center; justify-content: flex-end; height: 100%;';
     // var cellStyle = 'display: table-cell; vertical-align: middle;';   
     /*Valida si valor total (el que se muestra
       abajo de tabla) es negativo.*/
       if (valor < 0)          
          return '<div style="' + style + 'color:red;">$' + valor.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '</div>';
       else
          return '<div style="' + style + '">$' + valor.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + '</div>';
  }

  /*Pipe (martes 9 julio 2024; 11:42 AM):
    genera número aleatorio para secuencia
    ACH.*/
		function GenerarNumAleatorio(min, max)
  {
     var numeroAleatorio;
     //Genera número aleatorio entre dos números.
     return numeroAleatorio = Math.floor((max - min + 1) * Math.random()) + min;

		};

  /*Pipe (martes 9 julio 2024; 12:06 AM):
    genera letra aleatoria para secuencia
    ACH.*/
		function GenerarLetraAleatoria()
  {
     var min, max, letraAleatoria;
     /*Números del 65 al 90 (código ASCII de
       letras de A a la Z).*/
     min = 65;
     max = 90;
     //Genera número aleatorio entre dos números.
     var numeroAleatorio = Math.floor((max - min + 1) * Math.random()) + min;
     /*Obtiene letra aleatoria según código ASCII (el
       valor de las letras mayúsculas comienzan en A = 65).*/
     return letraAleatoria = String.fromCharCode(numeroAleatorio);
		};
	</script>
</body>
</html>
