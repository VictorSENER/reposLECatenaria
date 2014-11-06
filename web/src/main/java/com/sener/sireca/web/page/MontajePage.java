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

import com.sener.sireca.web.bean.MontajeRevision;
import com.sener.sireca.web.bean.MontajeVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.MontajeService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class MontajePage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    // Dialog components
    @Wire
    Button newMontaje;
    @Wire
    Button handOverVersion;
    @Wire
    Listbox versionListBox;
    @Wire
    Grid revisionList;
    @Wire
    Label currentVersion;

    // Version list
    ListModelList<MontajeVersion> versionListModel;

    // Revision list
    ListModelList<MontajeRevision> revisionListModel;

    // Session data
    HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();

    // Services
    ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
    MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");
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
                montajeService.deleteVersion(project, numVersion);

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
                                        montajeService.deleteRevision(project,
                                                numVersion, numRevision);

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
                                                                    "/montaje/");
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
                                                                    "/montaje/");
                                                        }
                                                    }
                                                });
                                    }

                                }
                                else
                                    // Redirect back
                                    Executions.getCurrent().sendRedirect(
                                            "/montaje/");
                            }
                        });
            }
            else
                Executions.getCurrent().sendRedirect("/montaje/");
        }

        List<MontajeVersion> montajeVerList = montajeService.getVersions(project);

        currentVersion.setValue("Versión Actual: "
                + montajeVerList.get(montajeVerList.size() - 1).getNumVersion());

        for (int i = 0; i < montajeVerList.size(); i++)
            montajeVerList.get(i).setModelList(
                    montajeService.getRevisions(montajeVerList.get(i)));

        versionListModel = new ListModelList<MontajeVersion>(montajeVerList);
        versionListBox.setModel(versionListModel);

    }

    @Listen("onClick = #newMontaje")
    public void doMontajeoAdd()
    {
        Executions.getCurrent().sendRedirect("/montaje/new/");
    }

    @Listen("onClick = #handOverVersion")
    public void doHandOverVersion()
    {
        montajeService.createVersion(project);
        Executions.getCurrent().sendRedirect("/montaje");
    }

}
