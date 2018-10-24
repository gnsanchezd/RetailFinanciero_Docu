var request = require('request');
var excel = require('excel4node');
var _status    = ""
		, _body    = ""
		, _headers = "";
var _sprints = "";


request({
  method: 'GET',
  url: 'https://api.clickup.com/api/v1/space/455736/project',
  headers: {
    'Authorization': 'pk_A33IPMXOU2UJLFEYWF0WSIYSYC7PTB47'
  }}, function (error, response, body) {

  _sprints = JSON.parse( body );
  _sprints = _sprints.projects[0].lists;

//  _sprints = body.
});


request({
  method: 'GET',
  url: 'https://api.clickup.com/api/v1/team/453821/task',
  headers: {
    'Authorization': 'pk_A33IPMXOU2UJLFEYWF0WSIYSYC7PTB47'
  }}, function (error, response, body) {
  _status   = response.statusCode;
	_body 		= JSON.parse( body );
	_headers	= response.headers;

	//console.log( body );
	var par = true;
	if ( _status === 200 )
	{
		var workbook = new excel.Workbook();
		var worksheet = workbook.addWorksheet('Tasks');
		var style = workbook.createStyle({
  			fill: {
				type: 'pattern',
    			fgColor: '#E8E8E8',
    			patternType: 'solid'
  			},
			font: {
   				color: '#000000',
   				size: 12,
 			}
		});

		var style_w = workbook.createStyle({
  			fill: {
				 type: 'pattern',
    			fgColor: '#FFFFFF',
    			patternType: 'solid'
  			},
			font: {
   				color: '#000000',
   				size: 12,
 			}
		});

		var style_cabecera = workbook.createStyle({
  			fill: {
				 type: 'pattern',
    			fgColor: '#175DA9',
    			patternType: 'solid'
  			},
			font: {
   				color: '#FFFFFF',
   				size: 12,
 			}
		});

		var style_fecha = workbook.createStyle({
  			fill: {
				 type: 'pattern',
    			fgColor: '#E20617',
    			patternType: 'solid'
  			},
			font: {
   				color: '#FFFFFF',
   				size: 12,
 			}
		});

		var _id_task = "";
		var _tasks = _body.tasks;
		var _name = "";
		var _text = "";
		var _status_task = "";
		var _creator_task = "";
		var _assigness_task = "";
		var _assigness_role = "";
		var _date_created = "";
		var _date_updated = "";
		var _date_closed = "";
		var _lists_id = "";
		var _lists_name = "";

		var separador = 3;

		worksheet.cell( 1, 1 ).string( "Fecha Sprint" ).style( style_fecha );
		worksheet.cell( 1, 2 ).string( _headers.date.toString() ).style( style_fecha );

		worksheet.cell( separador, 1 ).string( "ID Actividad" ).style( style_cabecera );
		worksheet.cell( separador, 2 ).string( "Actividad" ).style( style_cabecera );
		worksheet.cell( separador, 3 ).string( "Descripción" ).style( style_cabecera );
		worksheet.cell( separador, 4 ).string( "Estado" ).style( style_cabecera );
		worksheet.cell( separador, 5 ).string( "Sprint" ).style( style_cabecera );
		worksheet.cell( separador, 6 ).string( "Creador" ).style( style_cabecera );
		worksheet.cell( separador, 7 ).string( "Responsable" ).style( style_cabecera );

		for( var i = 0; i < _tasks.length ; i++ )
		{
			_id_task = _tasks[ i ].id;
			_name = _tasks[ i ].name;
			_text = _tasks[ i ].text_content;
			_creator_task = _tasks[ i ].creator.username;
			_status_task = _tasks[ i ].status.status.toUpperCase();
			_date_created = _tasks[ i ].date_created;
			_date_updated = _tasks[ i ].date_updated;
			_date_closed  = _tasks[ i ].date_closed;
			var _assigness_task = _tasks[ i ].assignees;

			for ( var  k = 0 ; k < _sprints.length ; k++ )
			{	
				if ( _sprints[ k ].id === _tasks[ i ].list.id )
				{
					_lists_name = _sprints[ k ].name;
				}
			}


			if ( par === true )
			{
				worksheet.cell( i + separador + 1, 1 ).string( validateNull( _id_task ) ).style( style );
				worksheet.cell( i + separador + 1, 2 ).string( validateNull(_name ) ).style( style );
				worksheet.cell( i + separador + 1, 3 ).string( validateNull(_text ) ).style( style );
				worksheet.cell( i + separador + 1, 4 ).string( validateNull(_status_task ) ).style( style );
				worksheet.cell( i + separador + 1, 5 ).string( validateNull(_lists_name ) ).style( style );
				worksheet.cell( i + separador + 1, 6 ).string( validateNull(_creator_task ) ).style( style );
			}
			else
			{
				worksheet.cell( i + separador + 1, 1 ).string( validateNull( _id_task ) ).style( style_w );
				worksheet.cell( i + separador + 1, 2 ).string( validateNull(_name ) ).style( style_w );
				worksheet.cell( i + separador + 1, 3 ).string( validateNull(_text ) ).style( style_w );
				worksheet.cell( i + separador + 1, 4 ).string( validateNull(_status_task ) ).style( style_w );
				worksheet.cell( i + separador + 1, 5 ).string( validateNull(_lists_name ) ).style( style_w );
				worksheet.cell( i + separador + 1, 6 ).string( validateNull(_creator_task ) ).style( style_w );
			}

			for ( var j = 0; j < _assigness_task.length ; j++ )
			{
				_assigness_role = _assigness_task[ j ].username;
				if ( par  === true )
				{
					worksheet.cell( i + separador + 1, j + 1 + 6 ).string( _assigness_role ).style( style );
				}
				else
				{
					worksheet.cell( i + separador + 1, j + 1 + 6 ).string( _assigness_role ).style( style_w );
				}

			}
			if ( par === true )
			{
				par = false;
			}
			else
			{
				par = true;
			}
		}
		var dateArchivo = _headers.date.toString( );
		dateArchivo = replaceAll( dateArchivo, ' ', '_' );
		dateArchivo = replaceAll( dateArchivo, ':', '_' );
		dateArchivo = replaceAll( dateArchivo, ',', '' );
		workbook.write('../../ProductBackLog/SnapShot/Task_RF_' + dateArchivo +'.xlsx');
	}
	else
	{
		console.log( "Error de conexión HTTP : " , _status );
	}
});


function validateNull( obj )
{
	if ( obj != null )
	{
		return obj.toString();
	}
	else
	{
		return "";
	}
}

function replaceAll(str, find, replace) {
    return str.replace(new RegExp(find, 'g'), replace);
}
