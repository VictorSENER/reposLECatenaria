/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.io.IOException;
import java.util.List;

import javax.servlet.http.HttpSession;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zul.Button;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Textbox;

import com.sener.sireca.web.bean.PendoladoRevision;
import com.sener.sireca.web.bean.PendoladoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.CatenariaService;
import com.sener.sireca.web.service.PendoladoService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.service.ReplanteoService;
import com.sener.sireca.web.util.SpringApplicationContext;
import com.sener.sireca.web.worker.PendoladoWorker;

public class PendoladoNewPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    // Dialog components
    @Wire
    Button fichaPendolado;
    @Wire
    Button volver;
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

    List<ReplanteoVersion> verList;

    // Session data
    HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();

    // Services
    ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
    ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");
    PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");
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

    }

    @SuppressWarnings({ "unchecked", "rawtypes" })
    @Listen("onChange = #versionList")
    public void fillRevisions()
    {
        List<Integer> rList = replanteoService.getRevisionList(verList.get(versionList.getSelectedIndex()));
        revisionList.setModel(new ListModelList(rList));
    }

    @Listen("onClick = #fichaPendolado; onOK=#pendoladoNewWin")
    public void doFichas() throws IOException
    {
        if (pkInicial.getValue().equals(""))
            Clients.showNotification("Debe introducir PK Inicial.");

        else if (pkFinal.getValue().equals(""))
            Clients.showNotification("Debe introducir PK Final.");

        else if (versionList.getValue().equals("Escoja Versión"))
            Clients.showNotification("Debe seleccionar una Version.");

        else if (revisionList.getValue().equals("Escoja Revisión"))
            Clients.showNotification("Debe seleccionar una Revisión.");

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
            int numVersion = pendoladoService.getLastVersion(project);
            int repVersion = versionList.getSelectedItem().getValue();
            int repRevision = revisionList.getSelectedItem().getValue();
            int idCatenaria = project.getIdCatenaria();

            String catenaria = catenariaService.getCatenariaById(idCatenaria).getNomCatenaria();

            ReplanteoVersion replanteoVersion = replanteoService.getVersion(
                    project, repVersion);
            ReplanteoRevision replanteoRevision = replanteoService.getRevision(
                    replanteoVersion, repRevision);

            PendoladoVersion pendoladoVersion = pendoladoService.getVersion(
                    project, numVersion);
            PendoladoRevision pendoladoRevision = pendoladoService.createRevision(
                    pendoladoVersion, replanteoRevision, notes.getValue());

            PendoladoWorker pw = new PendoladoWorker(pendoladoRevision, pkIni, pkFin, catenaria);

            pw.start();

            Executions.getCurrent().sendRedirect(
                    "/pendolado/progress/" + numVersion + "/"
                            + pendoladoRevision.getNumRevision());

        }

    }

    @Listen("onClick = #volver")
    public void doGoBack()
    {
        // Go back
        Executions.getCurrent().sendRedirect("/pendolado");
    }

}
