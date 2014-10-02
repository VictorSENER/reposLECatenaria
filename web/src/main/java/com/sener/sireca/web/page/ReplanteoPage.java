/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.io.FileInputStream;
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
import org.zkoss.zul.Button;
import org.zkoss.zul.Filedownload;
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
                                    try
                                    {

                                        if (!replanteoService.getRevision(
                                                replanteoService.getVersion(
                                                        project, numVersion),
                                                numRevision).getCalculated())
                                            throw new Exception();

                                        replanteoService.deleteRevision(
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
                                                                    "/replanteo/");
                                                        }

                                                    }
                                                });

                                    }
                                    catch (Exception e1)
                                    {

                                        Messagebox.show(
                                                "Fallo al eliminar la revisión "
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
                                                                    "/replanteo/");
                                                        }
                                                    }
                                                });
                                    }

                                }
                                else
                                    // Redirect back
                                    Executions.getCurrent().sendRedirect(
                                            "/replanteo/");
                            }
                        });
            }

            else if (action.equals("show"))
            {
                Messagebox.show("Revisión " + numRevision + " de la versión "
                        + numVersion + " importada correctamente.",
                        "Información", Messagebox.OK, Messagebox.INFORMATION,
                        new org.zkoss.zk.ui.event.EventListener<Event>()
                        {
                            @Override
                            public void onEvent(Event e) throws Exception
                            {

                                if (e.getName().equals("onOK"))
                                    // Redirect back
                                    Executions.getCurrent().sendRedirect(
                                            "/replanteo/");
                            }
                        });
            }
            else
                Executions.getCurrent().sendRedirect("/replanteo/");
        }

        List<ReplanteoVersion> replanteoVerList = replanteoService.getVersions(project);

        currentVersion.setValue("Versión Actual: "
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
        java.io.InputStream is = new FileInputStream(System.getenv("SIRECA_HOME")
                + "/templates/INPUTS_template.xlsx");
        Filedownload.save(is, "application/xlsx", "INPUTS_template.xlsx");

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
