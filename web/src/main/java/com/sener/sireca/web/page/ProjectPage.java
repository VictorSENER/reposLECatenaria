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
import org.zkoss.zk.ui.event.ForwardEvent;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zul.Button;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.Messagebox;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.service.ActiveProjectService;
import com.sener.sireca.web.service.AuthenticationService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class ProjectPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    // Dialog components
    @Wire
    Button addProject;
    @Wire
    Listbox projectListbox;
    @Wire
    Component selectedProjectBlock;

    // Projects list
    ListModelList<Project> projectListModel;

    // Project currently selected.
    Project selectedProject;

    // Session data
    HttpSession session = (HttpSession) Sessions.getCurrent().getNativeSession();
    AuthenticationService authService = (AuthenticationService) SpringApplicationContext.getBean("authService");
    ActiveProjectService actProj = (ActiveProjectService) SpringApplicationContext.getBean("actProj");
    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");

    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        // Fill projects list using DB data
        List<Project> projectList = projectService.getAllProjects();
        projectListModel = new ListModelList<Project>(projectList);
        projectListbox.setModel(projectListModel);

    }

    @Listen("onClick = #addproject")
    public void doProjectAdd()
    {

        // Get a title for the new project.
        String title = buildNewProjectTitle();

        // Create new project object.
        Project project = new Project();
        project.setTitulo(title);
        project.setIdUsuario(authService.getUserCredential(session).getIdUser());
        project.setCliente("Nombre Cliente");
        project.setReferencia("Referencia");

        // Store new project into DB.
        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
        projectService.insertProject(project);

        // Add new project into list model and select it.
        selectedProject = projectService.getProjectByTitle(title);
        projectListModel.add(selectedProject);
        projectListModel.addToSelection(selectedProject);

    }

    @Listen("onClick = #selectproject")
    public void doProjectSelectActive()
    {

        if (selectedProject != null)
        {
            actProj.setActive(session, selectedProject.getId(),
                    selectedProject.getTitulo());
            Executions.sendRedirect("/project");
        }

        else
            Messagebox.show("No ha seleccionado ningún proyecto.",
                    "Información", Messagebox.OK, Messagebox.INFORMATION);

    }

    @Listen("onProjectDelete = #projectListbox")
    public void doProjectDelete(final ForwardEvent evt)
    {
        // Ask for project confirmation.
        Messagebox.show("Está seguro que quiere borrar este proyecto?",
                "Confirmación", Messagebox.OK | Messagebox.CANCEL,
                Messagebox.QUESTION,
                new org.zkoss.zk.ui.event.EventListener<Event>()
                {
                    public void onEvent(Event e) throws Exception
                    {
                        if (e.getName().equals("onOK"))
                        {

                            // Get project to be deleted.
                            Button btn = (Button) evt.getOrigin().getTarget();
                            Listitem litem = (Listitem) btn.getParent().getParent();
                            Project project = (Project) litem.getValue();

                            // Delete project from DB.
                            ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
                            projectService.deleteProject(project.getId());

                            // Remove project from listbox.
                            projectListModel.remove(project);

                            // Refresh view when necessary.
                            if (project.equals(selectedProject))
                                selectedProject = null;

                            // Show confirmation.
                            Clients.showNotification("Proyecto borrado correctamente");

                        }
                    }
                });
    }

    @Listen("onProjectEdit = #projectListbox")
    public void doProjectEdit(final ForwardEvent evt)
    {

        Button btn = (Button) evt.getOrigin().getTarget();
        Listitem litem = (Listitem) btn.getParent().getParent();
        Project project = (Project) litem.getValue();

        Executions.getCurrent().sendRedirect("/project/edit/" + project.getId());

    }

    @Listen("onSelect = #projectListbox")
    public void doProjectSelect()
    {
        // Update selected project member
        if (projectListModel.isSelectionEmpty())
            selectedProject = null;
        else
            selectedProject = projectListModel.getSelection().iterator().next();

    }

    private String buildNewProjectTitle()
    {
        // Check if base project title isn't used.
        String baseProjecttitle = "Nuevo proyecto";
        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
        if (projectService.getProjectByTitle(baseProjecttitle) == null)
            return baseProjecttitle;

        int sequential = 1;
        while (sequential < 100)
        {
            String seqProjectname = baseProjecttitle + sequential;
            if (projectService.getProjectByTitle(seqProjectname) == null)
                return seqProjectname;

            sequential++;
        }

        return "nuevo";
    }
}
