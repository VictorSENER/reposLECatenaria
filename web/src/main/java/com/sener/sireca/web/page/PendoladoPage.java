/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.util.List;

import javax.servlet.http.HttpSession;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.Sessions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zul.Button;
import org.zkoss.zul.Grid;
import org.zkoss.zul.Label;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Messagebox;

import com.sener.sireca.web.bean.PendoladoRevision;
import com.sener.sireca.web.bean.PendoladoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.PendoladoService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class PendoladoPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    // Dialog components
    @Wire
    Button newPendolado;
    @Wire
    Button handOverVersion;
    @Wire
    Listbox versionListBox;
    @Wire
    Grid revisionList;
    @Wire
    Label currentVersion;

    // Version list
    ListModelList<PendoladoVersion> versionListModel;

    // Revision list
    ListModelList<PendoladoRevision> revisionListModel;

    // Session data
    HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();

    // Services
    ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
    PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");
    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

    final Project project = projectService.getProjectById(actProj.getIdActive(session));

    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        String action = (String) Executions.getCurrent().getAttribute("action");

        if (!action.equals(""))
        {

            final int numVersion = (Integer) Executions.getCurrent().getAttribute(
                    "numVersion");
            final int numRevision = (Integer) Executions.getCurrent().getAttribute(
                    "numRevision");

            if (action.equals("deleteVer"))
                pendoladoService.deleteVersion(project, numVersion);

            else if (action.equals("delete"))
            {
                Messagebox.show(
                        "Está seguro que quiere eliminar esta revisión?",
                        "Confirmación", Messagebox.OK | Messagebox.CANCEL,
                        Messagebox.QUESTION,
                        new org.zkoss.zk.ui.event.EventListener<Event>()
                        {
                            @Override
                            public void onEvent(Event e) throws Exception
                            {
                                if (e.getName().equals("onOK"))
                                {
                                    try
                                    {
                                        pendoladoService.deleteRevision(
                                                project, numVersion,
                                                numRevision);

                                        Messagebox.show(
                                                "Revisión "
                                                        + numRevision
                                                        + " de la versión "
                                                        + numVersion
                                                        + " eliminada correctamente.",
                                                "Información",
                                                Messagebox.OK,
                                                Messagebox.INFORMATION,
                                                new org.zkoss.zk.ui.event.EventListener<Event>()
                                                {
                                                    @Override
                                                    public void onEvent(Event e)
                                                            throws Exception
                                                    {

                                                        if (e.getName().equals(
                                                                "onOK"))
                                                        {
                                                            // Redirect back
                                                            Executions.getCurrent().sendRedirect(
                                                                    "/pendolado/");
                                                        }

                                                    }
                                                });
                                    }
                                    catch (Exception ex)
                                    {
                                        Messagebox.show(
                                                ex.getMessage(),
                                                "Información",
                                                Messagebox.OK,
                                                Messagebox.INFORMATION,
                                                new org.zkoss.zk.ui.event.EventListener<Event>()
                                                {
                                                    @Override
                                                    public void onEvent(Event e)
                                                    {
                                                        if (e.getName().equals(
                                                                "onOK"))
                                                        {
                                                            // Redirect back
                                                            Executions.getCurrent().sendRedirect(
                                                                    "/pendolado/");
                                                        }
                                                    }
                                                });

                                    }
                                }
                                else
                                    // Redirect back
                                    Executions.getCurrent().sendRedirect(
                                            "/pendolado/");
                            }
                        });
            }
            else
                Executions.getCurrent().sendRedirect("/pendolado/");
        }

        List<PendoladoVersion> pendoladoVerList = pendoladoService.getVersions(project);

        currentVersion.setValue("Versión Actual: "
                + pendoladoVerList.get(pendoladoVerList.size() - 1).getNumVersion());

        for (int i = 0; i < pendoladoVerList.size(); i++)
            pendoladoVerList.get(i).setModelList(
                    pendoladoService.getRevisions(pendoladoVerList.get(i)));

        versionListModel = new ListModelList<PendoladoVersion>(pendoladoVerList);
        versionListBox.setModel(versionListModel);

    }

    @Listen("onClick = #newPendolado")
    public void doPendoladoAdd()
    {
        Executions.getCurrent().sendRedirect("/pendolado/new/");
    }

    @Listen("onClick = #handOverVersion")
    public void doHandOverVersion()
    {
        pendoladoService.createVersion(project);
        Executions.getCurrent().sendRedirect("/pendolado");
    }

}
