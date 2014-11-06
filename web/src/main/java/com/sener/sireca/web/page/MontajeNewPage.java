/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.io.File;
import java.io.IOException;
import java.util.List;

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
import org.zkoss.zul.Combobox;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Textbox;

import com.sener.sireca.web.bean.MontajeRevision;
import com.sener.sireca.web.bean.MontajeVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.CatenariaService;
import com.sener.sireca.web.service.MontajeService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.service.ReplanteoService;
import com.sener.sireca.web.util.SpringApplicationContext;
import com.sener.sireca.web.worker.MontajeWorker;

public class MontajeNewPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    // Dialog components
    @Wire
    Button fichaMontaje;
    @Wire
    Button volver;
    @Wire
    Button uploadFile;
    @Wire
    Textbox fileToUpload;
    @Wire
    Textbox pkInicial;
    @Wire
    Textbox pkFinal;
    @Wire
    Textbox notes;
    @Wire
    Combobox versionList;
    @Wire
    Combobox revisionList;
    @Wire
    Checkbox autocad;
    @Wire
    Checkbox pdf;

    private List<ReplanteoVersion> verList;

    private Media media = null;

    // Session data
    HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();

    // Services
    ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
    ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");
    MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");
    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
    CatenariaService catenariaService = (CatenariaService) SpringApplicationContext.getBean("catenariaService");

    @SuppressWarnings({ "rawtypes", "unchecked" })
    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        Project project = projectService.getProjectById(actProj.getIdActive(session));

        verList = replanteoService.getVersions(project);

        List<Integer> vList = replanteoService.getVersionList(project);

        versionList.setModel(new ListModelList(vList));

        versionList.setValue("Escoja Versión");
        revisionList.setValue("Escoja Revisión");

        uploadFile.setUpload("true");

        uploadFile.addEventListener("onUpload",
                new EventListener<UploadEvent>()
                {
                    @Override
                    public void onEvent(UploadEvent event) throws Exception
                    {
                        media = event.getMedia();
                        String fileName = media.getName();
                        if (fileName.endsWith(".dwg"))
                            fileToUpload.setValue(fileName);
                        else
                            Clients.showNotification("Debe subir un fichero con el formato de la plantilla indicada.");
                    }

                });

    }

    @SuppressWarnings({ "unchecked", "rawtypes" })
    @Listen("onChange = #versionList")
    public void fillRevisions()
    {
        revisionList.setValue("Escoja Revisión");
        List<Integer> rList = replanteoService.getRevisionList(verList.get(versionList.getSelectedIndex()));
        revisionList.setModel(new ListModelList(rList));
    }

    @Listen("onClick = #fichaMontaje; onOK=#montajeNewWin")
    public void doFichas() throws IOException
    {

        if (fileToUpload.getValue().equals(""))
            Clients.showNotification("No ha seleccionado ningun archivo.");

        else if (pkInicial.getValue().equals(""))
            Clients.showNotification("Debe introducir PK Inicial.");

        else if (pkFinal.getValue().equals(""))
            Clients.showNotification("Debe introducir PK Final.");

        else if (versionList.getValue().equals("Escoja Versión"))
            Clients.showNotification("Debe seleccionar una Version.");

        else if (revisionList.getValue().equals("Escoja Revisión"))
            Clients.showNotification("Debe seleccionar una Revisión.");

        else if (!(autocad.isChecked() || pdf.isChecked()))
            Clients.showNotification("Debe seleccionar al menos un formato.");

        else
        {

            double pkIni = 0;
            double pkFin = 0;

            try
            {
                pkIni = Double.parseDouble(pkInicial.getValue().replace(',',
                        '.'));
                pkFin = Double.parseDouble(pkFinal.getValue().replace(',', '.'));
            }
            catch (Exception e)
            {
                Clients.showNotification("El PK debe ser un valor numérico.");
                return;
            }

            if (pkIni >= pkFin)
            {
                Clients.showNotification("El PK Inicial debe ser menor que el PK Final.");
                return;
            }

            Project project = projectService.getProjectById(actProj.getIdActive(session));
            int numVersion = montajeService.getLastVersion(project);
            int repVersion = versionList.getSelectedItem().getValue();
            int repRevision = revisionList.getSelectedItem().getValue();
            int idCatenaria = project.getIdCatenaria();
            boolean bCAD = autocad.isChecked();
            boolean bPDF = pdf.isChecked();

            String catenaria = catenariaService.getCatenariaById(idCatenaria).getNomCatenaria();

            ReplanteoVersion replanteoVersion = replanteoService.getVersion(
                    project, repVersion);
            ReplanteoRevision replanteoRevision = replanteoService.getRevision(
                    replanteoVersion, repRevision);

            MontajeVersion montajeVersion = montajeService.getVersion(project,
                    numVersion);
            MontajeRevision montajeRevision = montajeService.createRevision(
                    montajeVersion, replanteoRevision, notes.getValue());

            String ruta = montajeRevision.getAutoCadPath();

            File dest = new File(ruta);
            Files.copy(dest, media.getStreamData());

            MontajeWorker mw = new MontajeWorker(montajeRevision, pkIni, pkFin, catenaria, bPDF, bCAD);

            mw.start();

            Executions.getCurrent().sendRedirect(
                    "/montaje/progress/" + numVersion + "/"
                            + montajeRevision.getNumRevision());
        }

    }

    @Listen("onClick = #volver")
    public void doGoBack()
    {
        // Go back
        Executions.getCurrent().sendRedirect("/montaje");
    }

}
