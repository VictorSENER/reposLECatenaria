/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.io.File;
import java.io.IOException;

import javax.servlet.http.HttpSession;

import org.zkoss.io.Files;
import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.event.EventListener;
import org.zkoss.zk.ui.event.UploadEvent;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zul.Button;
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Textbox;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.CatenariaService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.service.ReplanteoService;
import com.sener.sireca.web.util.SpringApplicationContext;
import com.sener.sireca.web.worker.ReplanteoWorker;

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
    ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");
    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
    CatenariaService catenariaService = (CatenariaService) SpringApplicationContext.getBean("catenariaService");

    Media media = null;

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
                        media = event.getMedia();
                        String fileName = media.getName();
                        if (fileName.endsWith(".xlsx")
                                || fileName.endsWith(".xls"))
                            fileToUpload.setValue(fileName);
                        else
                            Clients.showNotification("Debe subir un fichero con el formato de la plantilla indicada.");
                    }

                });

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
        // Go back
        Executions.getCurrent().sendRedirect("/replanteo");
    }

    @Listen("onClick = #calculoReplanteo")
    public void doCalculateReplanteo() throws IOException
    {
        Project project = projectService.getProjectById(actProj.getIdActive(session));
        int numVersion = replanteoService.getLastVersion(project);
        ReplanteoVersion replanteoVersion = replanteoService.getVersion(
                project, numVersion);

        ReplanteoRevision replanteoRevision;

        if (fileToUpload.getValue().equals(""))
            Clients.showNotification("No ha seleccionado ningun archivo.");
        else
        {

            if (calcularImportar.isChecked())
            {
                replanteoRevision = replanteoService.createRevision(
                        replanteoVersion, 0);

                String ruta = replanteoRevision.getExcelPath();

                File dest = new File(ruta);
                Files.copy(dest, media.getStreamData());

                int idCatenaria = project.getIdCatenaria();
                long pkIni = Integer.parseInt(pkInicial.getValue());
                long pkFin = Integer.parseInt(pkFinal.getValue());

                String catenaria = catenariaService.getCatenariaById(
                        idCatenaria).getNomCatenaria();

                ReplanteoWorker rw = new ReplanteoWorker(replanteoRevision, pkIni, pkFin, catenaria);

                rw.start();

                // TODO: Calculate Revision
                Executions.getCurrent().sendRedirect(
                        "/replanteo/progress/" + numVersion + "/"
                                + replanteoRevision.getNumRevision());
            }
            else
            {
                replanteoRevision = replanteoService.createRevision(
                        replanteoVersion, 1);

                String ruta = replanteoRevision.getExcelPath();

                File dest = new File(ruta);
                Files.copy(dest, media.getStreamData());

                Executions.getCurrent().sendRedirect(
                        "/replanteo/show/" + numVersion + "/"
                                + replanteoRevision.getNumRevision());

            }
        }

    }
}
