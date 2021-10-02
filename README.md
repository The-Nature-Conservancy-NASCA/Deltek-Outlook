# Deltek-Outlook

## ¿Qué hace esta cosa?
Este código en Python permite estimar las horas dedicadas en los diferentes proyectos a partir de reuniones programadas en Outlook.

## ¿Cómo funciona esto?
Funcionalmente, el código se conecta a la aplicación de Outlook de escritorio (Windows) y realiza la lectura de todas las reuniones que se tengan programadas entre un intervalo de tiempo definido, luego realiza un filtrado por una palabra clave (en el asunto) y finalmente se agrupan las horas por una condición común.

## ¿Qué tengo que hacer para usarlo?
Lo primero que tienes que hacer para utilizar este código, es sistematizar el asunto de tus reuniones. Claramente, son muchas las reuniones que te llegan a diario y con múltiples asuntos. Probablemente algunas no se relacionen con tus proyectos, por lo que es necesario utilizar un sistema. El cogido desarrollado utiliza el siguiente sistema en el asunto de las reuniones:

**Deltek | Proyecto | Asunto**

La idea de este sistema es que el código mapee solamente las reuniones que inicien en su asunto con **Deltek**. Ahora bien, para saber a qué proyecto se debe asignar la hora, es necesario identificarlo y por ultimo se asignan el asunto. El código entiende la separación entre esto con un **|** por lo que es indispensable seguir esta estructura para su poder usarlo. Esto demandara de ti mucha diciplina para mantener tu agenda organizada. Sabiendo esto, lo único que necesitas hacer es tener instaladas las siguientes librerías en Python.

- Datetime
- Pandas
- Numpy
- win32com

Con esto y el IDE que quieras, puedes usar el código

## ¿Cómo la hago funcionar?
**¡¡¡Muy fácil!!!**, lo primero que tiene que hacer es en la línea 51 y 52 del código, definir la fecha de inicio y la fecha de finalización en la cual quieres realizar el mapeo. La definimos de la siguiente manera: Año, mes, día. Por ejemplo, si quisiera que el código mapeara las reuniones desde el 1 hasta el 31 de octubre, la configuración seria la siguiente:

	Start_Time  = dt.datetime(2021,10,1)
	End_Time    = dt.datetime(2021,10,31)

Luego tenemos que definir la palabra clave para que filtre las reuniones (en la línea 53). Para el caso particular de este código, se utilizó la palabra “Deltek”. Sin embargo, tu puede definir la palabra que quieras, siempre y cuando sistematices tu calendario con ella.

	keyword     = 'Deltek'

Con estas dos configuraciones, lo que reta es ejecutar y listo.

## ¿Muy chévere y todo, pero, cuáles son los resultados?
Ten calma, cuando ejecutes el código, este te arrojara 2 archivos Excel en la misma carpeta del código. Uno de ellos tiene el nombre de **01-Report.xlsx**, este contiene el registro de todas las reuniones y consolidad información como el proyecto al cual pertenece la reunión, hora de inicio de la reunión, hora de finalización, fecha, la descripción y las horas dedicadas. El segunto archivo **02-Deltek**, contiene el resumen de las horas dedicadas día a día a cada proyecto en el periodo de tiempo definió.
Hey, creo que esta información te va ha se muy útil y creara en ti un habito de disciplina para gestionar tus horas de trabajo en cada proyecto.
