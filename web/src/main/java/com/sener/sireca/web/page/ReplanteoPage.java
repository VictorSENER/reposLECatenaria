/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.io.FileNotFoundException;
import java.util.List;

import javax.servlet.http.HttpSession;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zul.Button;
import org.zkoss.zul.Grid;
import org.zkoss.zul.Label;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Messagebox;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoRevision;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.service.ReplanteoService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class ReplanteoPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    // Dialog components
    @Wire
    Button newReplanteo;
    @Wire
    Button handOverVersion;
    @Wire
    Button downloadTemplate;
    @Wire
    Listbox versionListBox;
    @Wire
    Grid revisionList;
    @Wire
    Label currentVersion;

    // Version list
    ListModelList<ReplanteoVersion> versionListModel;

    // Revision list
    ListModelList<ReplanteoRevision> revisionListModel;

    // Session data
    HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();

    // Services
    ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
    ReplanteoService replanteoService = (ReplanteoService) SpringApplicationContext.getBean("replanteoService");
    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

    final Project project = projectService.getProjectById(actProj.getIdActive(session));

    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        String action = (String) Executions.getCurrent().getAttribute("action");

        if (action.equals("delete"))
        {

            final int numVersion = (Integer) Executions.getCurrent().getAttribute(
                    "numVersion");
            final int numRevision = (Integer) Executions.getCurrent().getAttribute(
                    "numRevision");

            Messagebox.show("Está seguro que quiere eliminar esta revisión?",
                    "Confirmación", Messagebox.OK | Messagebox.CANCEL,
                    Messagebox.QUESTION,
                    new org.zkoss.zk.ui.event.EventListener<Event>()
                    {
                        public void onEvent(Event e) throws Exception
                        {

                            if (e.getName().equals("onOK"))
                            {

                                replanteoService.deleteRevision(project,
                                        numVersion, numRevision);

                                // Show confirmation.
                                Clients.showNotification("Revision eliminada correctamente"
                                        + numVersion + "_" + numRevision);

                            }
                            Executions.getCurrent().sendRedirect("/replanteo/");
                        }
                    });

        }

        List<ReplanteoVersion> replanteoVerList = replanteoService.getVersions(project);

        currentVersion.setValue("Version Actual: "
                + replanteoVerList.get(replanteoVerList.size() - 1).getNumVersion());

        for (int i = 0; i < replanteoVerList.size(); i++)
            replanteoVerList.get(i).setModelList(
                    replanteoService.getRevisions(replanteoVerList.get(i)));

        versionListModel = new ListModelList<ReplanteoVersion>(replanteoVerList);
        versionListBox.setModel(versionListModel);

    }

    @Listen("onClick = #downloadTemplate")
    public void doDownloadTemplate() throws FileNotFoundException
    {
        // TODO: descargar el template.

    }

    @Listen("onClick = #newReplanteo")
    public void doReplanteoAdd()
    {
        Executions.getCurrent().sendRedirect("/replanteo/new/");
    }

    @Listen("onClick = #handOverVersion")
    public void doHandOverVersion()
    {
        replanteoService.createVersion(project);
        Executions.getCurrent().sendRedirect("/replanteo");
    }

}
