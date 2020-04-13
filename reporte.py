"""
para visualizar la aplicación de riesgos en la que estoy trabajando
hace falta entrar a www.resiliente.org logearse con johandanielcruz@gmail.com 
y 123

para hacer el reporte que es el código que estoy compartiendo
hace falta tener plantilla_reporte.xlsx con el preformato
del documento

En esta app he trabajado con el orm sqlalchemy tiene valores de fábrica 
que se quiere que se copien cuando se crea una nueva empresa, por 
el momento se quiere probar con un solo cliente 

a la app le faltan varias cosas (no soy robot 3 version en el login,
mejoras en la visualización en celulares); yo trabajo como contratista 
y cada ciclo son nuevos retos

voy a copiar la ultima propuesta que hice a ese cliente, ellos son 
"expertos" en normatividad y tienen un socio con experiencia en riesgos

ellos me indican en reuniones que funciones quieren y yo traduzco a casos
de uso hago propuestas discutimos números, acordamos fechas y asi hemos 
avanzado

inicialmente los requerimientos fueron unos, luego esos mismos cambiaron
yo me di cuenta de errores en la info que me entregaba el "experto" en
riesgos asociado estaba llevando mal sus tablas y fué necesario
reunirnos y ayudar a corregir sus herramientas

ese experto entro en otros proyectos se ha extendido el tiempo en que
no he programado para finalizar la app, estamos en negociación

"""

@main_blueprint.route('/reporte', methods=["POST", "GET" ])
@main_blueprint.route('/reporte/<name>/<location>', methods=["POST", "GET" ])
@login_required
@empleado
def reporte(name="X",location="Y"):
	import openpyxl
	from openpyxl.utils import FORMULAE
	wb= openpyxl.load_workbook('data/plantilla_reporte.xlsx')
	sheet=wb.get_sheet_by_name('riesgos')

	

	if current_user.es_empleado:
		e = models.Empresa.query.filter_by(id=current_user.empresa_id).first()
	else:
		e = models.Empresa.query.filter_by(razonSocial="valores de fabrica").first()
	

	tipo_De_Riesgos=models.Tipoderiesgo.query.all()

	c=1
	id_sgsst=-1

	for t in tipo_De_Riesgos:
		if c==6:
			id_sgsst=t.id
		c += 1


	riesgos= models.Riesgo.query.filter((models.Riesgo.empresa_id==e.id) & ( models.Riesgo.tipoDeRiesgo_id != id_sgsst)).all()#order_by(models.Riesgo.tipoDeRiesgo_id.asc()).all()

	fila=12
	columna=["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U"]

	#diana

	t_r={}
	for tipo in tipo_De_Riesgos:
		t_r[str(tipo.id)]=tipo.nombre
	print("t_r",t_r)
	tipo_De_Riesgos=""

	for riesgo in riesgos:
		n=0
		subproceso=models.Subproceso.query.filter_by(id=riesgo.subproceso_id).first()
		if subproceso is not None:
			proceso =models.Proceso.query.filter_by(id=subproceso.proceso_id).first()
			if proceso is not None:
				sheet[columna[n]+str(fila)]=proceso.nombre
		n+=1
		if subproceso is not None:
			sheet[columna[n]+str(fila)]=subproceso.nombre
		n+=1		
		tipo_de_riesgo=str(t_r[str(riesgo.tipoDeRiesgo_id)])
		sheet[columna[n]+str(fila)]=tipo_de_riesgo
		
		n+=1
		sheet[columna[n]+str(fila)]=str(riesgo.nombre)

		n+=1
		sheet[columna[n]+str(fila)]=str(riesgo.descripcion)

		n+=1
		sheet[columna[n]+str(fila)]=str(riesgo.probabilidad)

		n+=1
		sheet[columna[n]+str(fila)]=str(riesgo.ifi+riesgo.ii+riesgo.il+riesgo.io)

		n+=1
		sheet[columna[n]+str(fila)]=str(riesgo.riesgo)

		n+=1
		#sheet[columna[n]+str(fila)]=str("zona de r Inherente")
		if riesgo.riesgo <=10:
			sheet[columna[n]+str(fila)]="Tolerable"
		elif riesgo.riesgo <=25:
			sheet[columna[n]+str(fila)]="Moderado"
		elif riesgo.riesgo ==30:
			if riesgo.probabilidad == 3:
				sheet[columna[n]+str(fila)]="Moderado"
			else:
				sheet[columna[n]+str(fila)]="Importante"
		
		elif riesgo.riesgo <=60:
			sheet[columna[n]+str(fila)]="Importante"

		elif riesgo.riesgo <=100:
			sheet[columna[n]+str(fila)]="Inaceptable"


		n+=1

		cadena_controles='='
		cadena_naturaleza='='
		contador=1
		#text='"hola"'
		controles = models.Control.query.filter(models.Control.riesgo_id==int(riesgo.id), models.Control.tipoDeControl <3).all()
		
		for control in controles:
			cadena_contador=str(contador)
			if control.descripcion != "Control..."  :
				cadena_controles += '"'+cadena_contador+ '. '+control.descripcion+ '"&CHAR(10)&'
				#cadena_controles ="="+text+'&CHAR(10)&'+text
				if control.tipoDeControl == 1:
						cadena_naturaleza += '"'+cadena_contador+ '. '+"Preventivo"+ '"&CHAR(10)&'
				elif control.tipoDeControl == 2:
					cadena_naturaleza += '"'+cadena_contador+ '. '+"Correctivo"+ '"&CHAR(10)&'
				
			contador +=1

		sheet[columna[n]+str(fila)]=str(cadena_controles[:len(cadena_controles) - 10] )
		
		

		n+=1
		sheet[columna[n]+str(fila)]=str(cadena_naturaleza[:len(cadena_naturaleza) - 10] )

		n+=1
		sheet[columna[n]+str(fila)]=str(riesgo.probabilidad_r)

		n+=1
		sheet[columna[n]+str(fila)]=str(riesgo.impactos_r)


		n+=1
		sheet[columna[n]+str(fila)]=str(riesgo.riesgo_r)

		n+=1
		if riesgo.riesgo_r <=10:
			sheet[columna[n]+str(fila)]="Tolerable"
		elif riesgo.riesgo_r <=25:
			sheet[columna[n]+str(fila)]="Moderado"
		elif riesgo.riesgo_r ==30:
			if riesgo.probabilidad_r == 3:
				sheet[columna[n]+str(fila)]="Moderado"
			else:
				sheet[columna[n]+str(fila)]="Importante"
		
		elif riesgo.riesgo_r <=60:
			sheet[columna[n]+str(fila)]="Importante"

		elif riesgo.riesgo_r <=100:
			sheet[columna[n]+str(fila)]="Inaceptable"





		n+=1
		sheet[columna[n]+str(fila)]=str("Acciones asociadas al control")


		n+=1
		sheet[columna[n]+str(fila)]=str("fecha de inicio")


		n+=1
		sheet[columna[n]+str(fila)]=str("fecha de terminación")

		n+=1
		sheet[columna[n]+str(fila)]=str("Responsable")


		n+=1
		if riesgo.riesgo!=0:
			sheet[columna[n]+str(fila)]="1-100*Rr/Ri = "+str((1-riesgo.riesgo_r/riesgo.riesgo)*100)+"%"
		else:
			sheet[columna[n]+str(fila)]="1-100*Rr/Ri = Indeterminado Riesgo Inherente es cero"
		n+=1
		sheet[columna[n]+str(fila)]=str("Evidencia")


 		




		fila+=1

	import random
	url="data/-reporte-"
	for n in range(0,10):
		url= url+random.choice(["a","b","g","A","B","G","1","2","3","4","5","6","7","8","9","0"])

	url= url+'.xlsx'

	#sheet['A12']="ADMINISTRACIÓN DE TECNOLOGÍAS E INFORMACIÓN hh"
	wb.save(url)
	


	return render_template('reporte.html',url=url)