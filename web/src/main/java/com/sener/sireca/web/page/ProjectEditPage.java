/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import java.util.List;

import org.zkoss.lang.Strings;
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
    Project selectedProject;

    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
    CatenariaService catenariaService = (CatenariaService) SpringApplicationContext.getBean("catenariaService");
    UserService userService = (UserService) SpringApplicationContext.getBean("userService");

    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        List<String> vList = catenariaService.getListCatenarias();

        selectedProjectCatenaria.setModel(new ListModelList(vList));

        int id = (Integer) Executions.getCurrent().getAttribute("id");
        selectedProject = projectService.getProjectById(id);

        selectedProjectTitle.setValue(selectedProject.getTitulo());
        selectedProjectClient.setValue(selectedProject.getCliente());
        selectedProjectUser.setValue(userService.getUserById(
                (selectedProject.getIdUsuario())).getUsername());
        selectedProjectReference.setValue(selectedProject.getReferencia());

        selectedProjectCatenaria.setValue(catenariaService.getCatenariaById(
                selectedProject.getIdCatenaria()).getNomCatenaria());

    }

    @Listen("onClick = #updateSelectedProject")
    public void doUpdateClick() throws InterruptedException
    {
        // Check if title is empty.
        if (Strings.isBlank(selectedProjectTitle.getValue()))
        {
            Clients.showNotification(
                    "El título del proyecto no puede estar vacío.",
                    selectedProjectTitle);
            return;
        }
        // Check if title is too long.
        else if (selectedProjectTitle.getValue().length() > 100)
        {
            Clients.showNotification(
                    "El título del proyecto no puede ser tan largo. (Máximo 100 carácteres)",
                    selectedProjectTitle);
            return;
        }

        // Check if client name is empty.
        if (Strings.isBlank(selectedProjectClient.getValue()))
        {
            Clients.showNotification(
                    "El nombre del cliente no puede estar vacío.",
                    selectedProjectClient);
            return;
        }

        // Check if reference is too long.
        else if (selectedProjectClient.getValue().length() > 50)
        {
            Clients.showNotification(
                    "El nombre del cliente no puede ser tan largo. (Máximo 50 carácteres)",
                    selectedProjectClient);
            return;
        }

        // Check if reference is empty.
        if (Strings.isBlank(selectedProjectReference.getValue()))
        {
            Clients.showNotification("La referencia no puede estar vacía.",
                    selectedProjectReference);
            return;
        }
        // Check if reference is too long.
        else if (selectedProjectReference.getValue().length() > 20)
        {
            Clients.showNotification(
                    "La referencia no puede ser tan larga. (Máximo 20 carácteres)",
                    selectedProjectReference);
            return;
        }

        // Set new data to selected user.
        selectedProject.setTitulo(selectedProjectTitle.getValue());
        selectedProject.setCliente(selectedProjectClient.getValue());
        selectedProject.setReferencia(selectedProjectReference.getValue());
        selectedProject.setIdCatenaria(catenariaService.getCatenariaByTitle(
                selectedProjectCatenaria.getSelectedItem().getValue().toString()).getId());

        // Save new data into DB.
        projectService.updateProject(selectedProject);

        // how message for user.
        Clients.showNotification("Proyecto guardado correctamente");

    }

    @Listen("onClick = #cancelSelectedProject")
    public void doCancelClick()
    {

        Messagebox.show("Está seguro que quiere volver?", "Confirmación",
                Messagebox.OK | Messagebox.CANCEL, Messagebox.QUESTION,
                new org.zkoss.zk.ui.event.EventListener<Event>()
                {
                    public void onEvent(Event e) throws Exception
                    {
                        if (e.getName().equals("onOK"))
                        {
                            selectedProject = null;

                            // Go back
                            Executions.getCurrent().sendRedirect("/project");
                        }
                    }
                });

    }
}
