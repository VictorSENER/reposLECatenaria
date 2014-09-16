/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import javax.servlet.http.HttpSession;

import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zul.Button;
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Textbox;

import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class ReplanteoNewPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    // Dialog components
    @Wire
    Button uploadFile;
    @Wire
    Button calculoReplanteo;
    @Wire
    Button volver;
    @Wire
    Textbox fileToUpload;
    @Wire
    Textbox pkInicial;
    @Wire
    Textbox pkFinal;
    @Wire
    Checkbox calcularImportar;

    // Session data
    HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();
    ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");

    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        uploadFile.setUpload("true");

        uploadFile.addEventListener("onUpload",
                new EventListener<UploadEvent>()
                {
                    @Override
                    public void onEvent(UploadEvent event) throws Exception
                    {
                        try
                        {
                            Media media = event.getMedia();

                            Clients.showNotification("upload details: "
                                    + " name "
                                    + media.getName()
                                    + " size "
                                    + (media.isBinary() ? media.getByteData().length
                                            : media.getStringData().length())
                                    + " type " + media.getContentType());
                        }
                        catch (Exception e)
                        {
                            e.printStackTrace();
                            Messagebox.show("Upload failed");
                        }
                    }

                });

    }

    // @Listen("onClick = #uploadFile")
    public void upload()
    {

        /*
         * org.zkoss.util.media.Media media = Fileupload.get();
         * 
         * File myFile = new File(media.getName());
         * 
         * fileToUpload.setText(media.getName());
         * 
         * UploadEvent event = (UploadEvent) ctx.getTriggerEvent(); Media media
         * = event.getMedia();
         * 
         * try { OutputStream outputStream = new FileOutputStream(new
         * File("../PRUEBA.xlsx")); InputStream inputStream =
         * media.getStreamData(); byte[] buffer = new byte[1024]; for (int
         * count; (count = inputStream.read(buffer)) != -1;) {
         * outputStream.write(buffer, 0, count); } outputStream.flush();
         * outputStream.close(); inputStream.close(); } catch (Exception e) {
         * fileToUpload.setText("MAL"); }
         */

    }

    @Listen("onCheck = #calcularImportar")
    public void changeSubmitStatus()
    {
        if (calcularImportar.isChecked())
        {
            pkInicial.setDisabled(false);
            pkFinal.setDisabled(false);
        }
        else
        {
            pkInicial.setDisabled(true);
            pkFinal.setDisabled(true);
        }
    }

    @Listen("onClick = #volver")
    public void doGoBack()
    {

        Messagebox.show("Está seguro que quiere volver?", "Confirmación",
                Messagebox.OK | Messagebox.CANCEL, Messagebox.QUESTION,
                new org.zkoss.zk.ui.event.EventListener<Event>()
                {
                    public void onEvent(Event e) throws Exception
                    {
                        if (e.getName().equals("onOK"))
                        {
                            // Go back
                            Executions.getCurrent().sendRedirect("/replanteo");
                        }
                    }
                });

    }

    @Listen("onClick = #calculoReplanteo")
    public void doCalculateReplanteo()
    {
        // TODO: Obtener la información de todo y mandarla.

        Executions.getCurrent().sendRedirect("/replanteo/progress/1/1");
    }

}
