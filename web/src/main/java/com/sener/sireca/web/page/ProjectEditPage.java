/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.util.List;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.event.Event;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Listen;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zk.ui.util.Clients;
import org.zkoss.zul.Button;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Textbox;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.service.CatenariaService;
import com.sener.sireca.web.service.ProjectService;
import com.sener.sireca.web.service.UserService;
import com.sener.sireca.web.util.SpringApplicationContext;

public class ProjectEditPage extends SelectorComposer<Component>
{
    private static final long serialVersionUID = 1L;

    // Dialog components

    @Wire
    Textbox selectedProjectTitle;
    @Wire
    Textbox selectedProjectUser;
    @Wire
    Textbox selectedProjectClient;
    @Wire
    Textbox selectedProjectReference;
    @Wire
    Button updateSelectedProject;
    @Wire
    Combobox selectedProjectCatenaria;

    // Projects list
    ListModelList<Project> projectListModel;

    // Project currently selected.
    Project project;

    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
    CatenariaService catenariaService = (CatenariaService) SpringApplicationContext.getBean("catenariaService");
    UserService userService = (UserService) SpringApplicationContext.getBean("userService");

    @SuppressWarnings({ "unchecked", "rawtypes" })
    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        List<String> vList = catenariaService.getListCatenarias();

        selectedProjectCatenaria.setModel(new ListModelList(vList));

        int id = (Integer) Executions.getCurrent().getAttribute("id");
        project = projectService.getProjectById(id);

        if (project == null)
            Executions.getCurrent().sendRedirect("/project");

        selectedProjectTitle.setValue(project.getTitulo());
        selectedProjectClient.setValue(project.getCliente());
        selectedProjectUser.setValue(userService.getUserById(
                (project.getIdUsuario())).getUsername());
        selectedProjectReference.setValue(project.getReferencia());

        selectedProjectCatenaria.setValue(catenariaService.getCatenariaById(
                project.getIdCatenaria()).getNomCatenaria());

    }

    @Listen("onClick = #updateSelectedProject")
    public void doUpdateClick() throws InterruptedException
    {

        // Set new data to selected user.
        project.setTitulo(selectedProjectTitle.getValue());
        project.setCliente(selectedProjectClient.getValue());
        project.setReferencia(selectedProjectReference.getValue());
        project.setIdCatenaria(catenariaService.getCatenariaByTitle(
                selectedProjectCatenaria.getSelectedItem().getValue().toString()).getId());

        // Save new data into DB.
        if (projectService.updateProject(project) != 0)
            // Show message for user.
            Clients.showNotification("Proyecto guardado correctamente");

    }

    @Listen("onClick = #cancelSelectedProject")
    public void doCancelClick()
    {

        Messagebox.show("Est� seguro que quiere volver?", "Confirmaci�n",
                Messagebox.OK | Messagebox.CANCEL, Messagebox.QUESTION,
                new org.zkoss.zk.ui.event.EventListener<Event>()
                {
                    @Override
                    public void onEvent(Event e) throws Exception
                    {
                        if (e.getName().equals("onOK"))
                            // Go back
                            Executions.getCurrent().sendRedirect("/project");

                    }
                });

    }
}
