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

import com.sener.sireca.web.bean.DibujoRevision;
import com.sener.sireca.web.bean.DibujoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.DibujoService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class DibujoPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    // Dialog components
    @Wire
    Button newDibujo;
    @Wire
    Button handOverVersion;
    @Wire
    Listbox versionListBox;
    @Wire
    Grid revisionList;
    @Wire
    Label currentVersion;

    // Version list
    ListModelList<DibujoVersion> versionListModel;

    // Revision list
    ListModelList<DibujoRevision> revisionListModel;

    // Session data
    HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();

    // Services
    ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
    DibujoService dibujoService = (DibujoService) SpringApplicationContext.getBean("dibujoService");
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

            if (action.equals("delete"))
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

                                    if (dibujoService.deleteRevision(project,
                                            numVersion, numRevision))

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
                                                                    "/drawing/");
                                                        }

                                                    }
                                                });

                                    else

                                        Messagebox.show(
                                                "Fallo al eliminar la revisión "
                                                        + numRevision
                                                        + " de la versión "
                                                        + numVersion + ".",
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
                                                                    "/drawing/");
                                                        }
                                                    }
                                                });

                                }
                                else
                                    // Redirect back
                                    Executions.getCurrent().sendRedirect(
                                            "/drawing/");
                            }
                        });
            }
            else
                Executions.getCurrent().sendRedirect("/drawing/");
        }

        List<DibujoVersion> dibujoVerList = dibujoService.getVersions(project);

        currentVersion.setValue("Versión Actual: "
                + dibujoVerList.get(dibujoVerList.size() - 1).getNumVersion());

        for (int i = 0; i < dibujoVerList.size(); i++)
            dibujoVerList.get(i).setModelList(
                    dibujoService.getRevisions(dibujoVerList.get(i)));

        versionListModel = new ListModelList<DibujoVersion>(dibujoVerList);
        versionListBox.setModel(versionListModel);

    }

    @Listen("onClick = #newDibujo")
    public void doDibujoAdd()
    {
        Executions.getCurrent().sendRedirect("/drawing/new/");
    }

    @Listen("onClick = #handOverVersion")
    public void doHandOverVersion()
    {
        dibujoService.createVersion(project);
        Executions.getCurrent().sendRedirect("/drawing");
    }

}
