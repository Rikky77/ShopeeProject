package objectRepository;

import org.w3c.dom.Document;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;

public class readObjectRepository {
    Document document;

    public Document read(String pageName) throws ParserConfigurationException, IOException, SAXException {
        if(pageName.trim().length() > 1) { //if page name is not empty
            File file = new File(System.getProperty("user.dir") + "\\src\\xmlObjectRepository\\" + pageName + ".xml");
            DocumentBuilderFactory builderFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder documentBuilder = builderFactory.newDocumentBuilder();

            document = documentBuilder.parse(file);
            document.getDocumentElement().normalize();

            return document;
        }
        else {
            return null;
        }
    }
}
