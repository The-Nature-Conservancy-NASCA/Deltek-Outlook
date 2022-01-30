# Deltek-Outlook

## ¿Qué hace esta cosa?
Basicamente 2 cosas. 1) estima las horas dedicadas en los diferentes proyectos a partir de reuniones programadas en Outlook. 2) diligencia de manera automatica el Deltek haciendo uso de Selenium-Python

## ¿Cómo funciona?
Fundamentalmente, el código **Deltek_Outlook.py** se conecta a la aplicación de Outlook de escritorio (Windows) y realiza la lectura de todas las reuniones que se tengan programadas entre un intervalo de tiempo definido, luego realiza un filtrado por una palabra clave (en el asunto) y finalmente se agrupan las horas por una condición común.

El código **Fill_Deltek.py** toma el csv **02-Deltek.csv** que genera el código **Deltek_Outlook.py** y el xlsx **00-Projects.xlsx** que debe ser condigurado con los Projects-ID, Activity y Adware de los proyectos en los cuales se está trabajando. Con esta información, el código diligencia de manera automática la hoja de tiempos de Deltek. **Pero OJO jejeje no todo es tan bueno**. Cuando tengas vacaciones o días festivos te tocara llenarlo a mano. Pero bueno ya tiene un avance jjejeje

## ¿Qué tengo que hacer para usarlo?
Lo primero que tienes que hacer para utilizar este conjunto de códigos, es sistematizar el asunto de tus reuniones. Claramente, son muchas las reuniones que te llegan a diario y con múltiples asuntos. Probablemente algunas no se relacionen con tus proyectos, por lo que es necesario utilizar un sistema. El cógido **Deltek_Outlook.py** utiliza el siguiente sistema en el asunto de las reuniones:

**Deltek | Proyecto | Asunto**

La idea de este sistema es que el código mapee solamente las reuniones que inicien en su asunto con **Deltek**. Ahora bien, para saber a qué proyecto se debe asignar la hora, es necesario identificarlo y por último se asignan el asunto. El código entiende la separación entre esto con un **|** por lo que es indispensable seguir esta estructura para su poder usarlo. Esto demandara de ti mucha diciplina para mantener tu agenda organizada. Sabiendo esto, lo único que necesitas hacer es tener instaladas las siguientes librerías en Python.

- Datetime
- Pandas
- Numpy
- win32com

Con esto y el IDE que quieras, puedes usar el código

Ahora bien, al código **Fill_Deltek.py** solo debes proporcionarle el número de empleado con el cual ingresas al Deltek y la contrasela.

## ¿Cómo la hago funcionar?
**¡¡¡Muy fácil!!!**, lo primero que tiene que hacer es en la línea 51 y 52 del código, definir la fecha de inicio y la fecha de finalización en la cual quieres realizar el mapeo. La definimos de la siguiente manera: Año, mes, día. Por ejemplo, si quisiera que el código mapeara las reuniones desde el 1 hasta el 31 de octubre, la configuración seria la siguiente:

	Start_Time  = dt.datetime(2021,10,1)
	End_Time    = dt.datetime(2021,10,31)

Luego tenemos que definir la palabra clave para que filtre las reuniones (en la línea 53). Para el caso particular de este código, se utilizó la palabra “Deltek”. Sin embargo, tu puede definir la palabra que quieras, siempre y cuando sistematices tu calendario con ella.

	keyword     = 'Deltek'

Con estas dos configuraciones, lo que resta es ejecutar **Deltek_Outlook.py**. Luego, debes ejecutar el código **Fill_Deltek.py** y listo, verifica que todo este correcto, guarda y firma tu hoja de tiempos.

## ¿Muy chévere y todo, pero, cuáles son las reglas?
Que creíste, que todo era color de rosas, pues no jejejeje. Este código debe ser ejecutado antes de que se acabe el mes que vas a reportar. Aún no está configurado para ubicar. Pero bueno, en próximas versiones seguro lo realizará. 
Aspectos claves, el código **Deltek_Outlook.py** te entregará un archivo que tiene por nombre **01-Report.xlsx**, este contiene el registro de todas las reuniones y consolidad información como el proyecto al cual pertenece la reunión, hora de inicio de la reunión, hora de finalización, fecha, la descripción y las horas dedicadas.
Otro de los archivos que tendras será **02-Deltek**, el cual contiene el resumen de las horas dedicadas día a día a cada proyecto en el periodo de tiempo definió. 
Por último también tendrás a tu disposición el archivo **03-Total_Deltek.csv** el contiene el resumen de horas totales trabajadas en cada proyecto.
Hey, creo que esta información te va ha se muy útil y creara en ti un habito de disciplina para gestionar tus horas de trabajo en cada proyecto.
