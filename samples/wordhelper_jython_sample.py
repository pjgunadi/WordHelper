import com.ibm.custom.WordHelper as WordHelper
import java.util.Calendar as Calendar
import java.text.DateFormat as DateFormat
import java.util.List as List
import java.util.Arrays as Arrays

df = DateFormat.getInstance()
rp = WordHelper("./samplet.dotx") #Create Helper Object
doc = rp.getDocument() #Sample method to get Document object

#Replace Keywords:
rp.replaceText("##LetterNum##", "123/456/789")
rp.replaceText("##CurrentDate##", df.format(Calendar.getInstance().getTime()))
rp.replaceText("##CustomerName##", "IBM")
rp.replaceText("##TicketID##", "ABC123")
rp.replaceText("##ResolveDate##", df.format(Calendar.getInstance().getTime()))
rp.replaceText("##SolutionDetails##", "Restart Server\n Recreate User\n")
rp.replaceText("##Resolver##", "John Doe")

#Update Tables - template table first row must match the header
tbheads = Arrays.asList("Circuit ID","Speed","Location")
tbdata = Arrays.asList(Arrays.asList("CI100","100 Mbps","Singapore"),Arrays.asList("CI210","1 Gbps","Kuala Lumpur"),Arrays.asList("CI320","10 Gbps","Bangkok"))
rp.updateTable(tbheads,tbdata)

#Save the result
rp.saveAs("./jython_output.doc")
