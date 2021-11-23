package com.example.pdfDocParser.controller;

import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.ContentAccessor;
import org.docx4j.wml.Text;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.FileSystemResource;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.mail.internet.MimeMessage;
import javax.xml.bind.JAXBElement;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Controller
@RequestMapping("/api")
public class MainController {

    @Autowired
    private JavaMailSender javaMailSender;

    @Value("${spring.mail.username}")
    private String mailFrom;

    public void sendEmail(String mail) throws Exception {
        MimeMessage message = javaMailSender.createMimeMessage();
        MimeMessageHelper helper = new MimeMessageHelper(message, true);
        helper.setTo(mail);
        helper.setSubject("Документ");
        helper.setText("Document");
        helper.setFrom(mailFrom);
        FileSystemResource file  = new FileSystemResource(new File("docx/templateResult.docx"));
        helper.addAttachment("templateResult.docx", file);
        javaMailSender.send(message);
    }

    @GetMapping
    public String getPage() {
        return "index";
    }


    @PostMapping
    @RequestMapping("/change")
    public String getDocumentInfo(@RequestBody String json) throws Exception {
        String email = "";
        try {
            Type mapType = new TypeToken<Map<String, String>>() {
            }.getType();
            Map<String, String> son = new Gson().fromJson(json, mapType);
            var template = getTemplate("docx/template.docx");
            for (Map.Entry<String, String> name : son.entrySet()) {
                String placeholder = name.getKey();
                String toAdd = name.getValue();
                if (placeholder.equals("email")) {
                    email = toAdd;
                }
                List<Object> texts = getAllElementFromObject(template.getMainDocumentPart(), Text.class);
                for (Object text : texts) {
                    Text textElement = (Text) text;
                    if (textElement.getValue().contains(placeholder)) {
                        String value = textElement.getValue().replace(placeholder, toAdd);
                        textElement.setValue(value);
                    }
                }
            }
            writeDocxToStream(template, "docx/templateResult.docx");
            sendEmail(email);
        } catch (Exception ignored) {
        }
        return "index";
    }

    private WordprocessingMLPackage getTemplate(String name) throws Docx4JException, FileNotFoundException {
        return WordprocessingMLPackage.load(new FileInputStream(new File(name)));
    }

    private static List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<Object>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        if (obj.getClass().equals(toSearch))
            result.add(obj);
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }

        }
        return result;
    }


    private void writeDocxToStream(WordprocessingMLPackage template, String target) throws Docx4JException {
        File f = new File(target);
        template.save(f);
    }
}
