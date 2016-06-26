# FronBot
Bot de entrenamiento inteligente basado en Argentum Online.

15/05/2015
-Repositorio iniciado.

21/05/2015
-Agregada Version 0.1.8.

Cambios:
*Agregada clase Guerrero con su IA.
*Agregado arco de cazador para el guerrero.
*Agregada funcion para poder equipar diferentes items.
*Agregados laud elfico y anillo de disolucion magica para el bardo.
*Agregados intervalos para los bots al remover a otro o resucitar a otro.
*Arreglos generales a varios sistemas e IAs.

20/05/2016
-Agregada Version 0.1.9.

Cambios:
*Agregadas teclas configurables para equipar, hablar y pausar.
*Optimización y limpieza del código.

24/05/2016
-Agregada Version 0.1.10.

Cambios:
*Agregadas opciones para elegir si se puede usar resucitar y si saca vida.
*Agregados codigos al .rar porque me tira error el git y no tengo tiempo ni ganas para ver por que.

25/05/2016
-Agregada Version 0.1.11.

Cambios:
*Ahora resucitar deja al objetivo sin mana.
*Agregadas estadisticas de falla y acierto al arco.
*Arreglado error que causaba que al lanzarle remover a un bot que estaba atacando al usuario inmo le provocaba comportarse como si hubiera sido atacado.
*Agregado sistema de prioridades segun clase y raza para los bots a la hora de seleccionar un objetivo.
*Arreglado error donde los bots se quedaban en fila tratando de atacar a un target inmo y mejorado el sistema para hacerlo mas fluido.

27/05/2015
-Agregada Version 0.1.12.

Cambios:
*Ahora los items no muestran cantidad en el inventario.
*Arreglado error que causaba que a los bots al lanzar resucitar les sacara mana y dijeran palabras magicas aunque no tuvieran la vida para lanzarlo.
*Agregado sistema de energia.
*Adaptada levemente la IA para el nuevo sistema de energia.
*Arreglado error que causaba que los bots pudieran remover y revivir fuera de su rango.
*Ahora los bot cuando mueren se mueven hacia quien los pueda revivir para no quedar fuera de su rango.
*Reducida la probabilidad del mago de atacar en vez de resucitar o remover a un compañero de 66,67% a 33,37%.

24/06/2016
-Agregada Version 0.2.0.

Cambios:
*Agregado modo online.
*Reconfigurado practicamente todo el bot para adaptarse al modo online.
*Agregada posibilidad de crear y unirse a partidas con y/o contra bots y/u otra gente.
*Agregado protocolo para cliente y servidor dentro del bot.
*Nuevo sistema de seleccion de equipo y personaje.
*Nuevo sistema de configuracion de partida y bots.

24/06/2016
-Agregada Version 0.2.1.

Cambios:
*Arreglado un error en el cierre del socket que causaba que no se limpie.
*Arreglado error que causaba runtime al recibir un mensaje de error como que el nombre de usuario ya estaba en uso al logear a causa de una llamada a mostrar un formulario con el mensaje de error abierto.
*Retiro lo dicho, modificado el error, ahora no tira runtime pero tampoco tira el mensaje de error, odio visual basic.
*Ahora cuando un usuario habla lo envia por consola a todos los usuarios.
*Arreglado error que impedia a un usuario conectarse si habia 9 personajes siendo usados.

26/06/2016
-Agregada Version 0.2.2.

Cambios:
*Cambiado el boton de cambiar personaje por un label para que no gane el foco involuntariamente, esto es temporal hasta meter el boton en la interfaz directamente.
*Arreglado bug que creaba un char en un lugar invalido al logear.
*Arreglado error que causaba que al deslogear un usuario los bots se bugearan.
*Arreglado un error que al resucitar no enviaba una actualizacion de personaje, causando que se siguiera viendo muerto.
*Agregada la L para "deslagear" la posicion.

26/06/2016
-Agregada Version 0.2.3.

Cambios:
*Agregados textos faltantes al recibir un hechizo de un usuario.
*Arreglado error que al intentarse mover en una posicion invalida no actualizaba la posicion correcta.
*Arreglada la L que no funcionaba del todo bien.
*Arreglados los errores de pisar a otros usuarios y que impedian moverse en a una posicion vacia.
