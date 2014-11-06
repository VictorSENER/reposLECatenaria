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
import org.zkoss.zul.Checkbox;
import org.zkoss.zul.Combobox;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Messagebox;
import org.zkoss.zul.Textbox;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.service.CatenariaService;
import com.sener.sireca.web.service.MontajeService;
import com.sener.sireca.web.service.PendoladoService;
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
    @Wire
    Combobox selectedPendoladoTemplate;
    @Wire
    Combobox selectedMontajeTemplate;
    @Wire
    Checkbox viaDoble;

    // Projects list
    ListModelList<Project> projectListModel;

    // Project currently selected.
    Project project;

    ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
    CatenariaService catenariaService = (CatenariaService) SpringApplicationContext.getBean("catenariaService");
    PendoladoService pendoladoService = (PendoladoService) SpringApplicationContext.getBean("pendoladoService");
    MontajeService montajeService = (MontajeService) SpringApplicationContext.getBean("montajeService");
    UserService userService = (UserService) SpringApplicationContext.getBean("userService");

    @SuppressWarnings({ "unchecked", "rawtypes" })
    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        int id = (Integer) Executions.getCurrent().getAttribute("id");
        project = projectService.getProjectById(id);

        if (project == null)
            Executions.getCurrent().sendRedirect("/project");
        List<String> vList = catenariaService.getListCatenarias();
        List<String> tPList = pendoladoService.getTemplatesList(project);
        List<String> tMList = montajeService.getTemplatesList(project);

        selectedProjectCatenaria.setModel(new ListModelList(vList));
        selectedPendoladoTemplate.setModel(new ListModelList(tPList));
        selectedMontajeTemplate.setModel(new ListModelList(tMList));

        selectedProjectTitle.setValue(project.getTitulo());
        selectedProjectClient.setValue(project.getCliente());
        selectedProjectUser.setValue(userService.getUserById(
                (project.getIdUsuario())).getUsername());

        selectedProjectReference.setValue(project.getReferencia());
        selectedProjectCatenaria.setValue(catenariaService.getCatenariaById(
                project.getIdCatenaria()).getNomCatenaria());

        selectedPendoladoTemplate.setValue(project.getPendolado());
        selectedMontajeTemplate.setValue(project.getMontaje());
        viaDoble.setChecked(project.getViaDoble());

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

        project.setPendolado(selectedPendoladoTemplate.getValue().toString());
        project.setMontaje(selectedMontajeTemplate.getValue().toString());
        project.setViaDoble(viaDoble.isChecked());

        try
        {

            projectService.updateProject(project);
            Clients.showNotification("Proyecto guardado correctamente");
        }
        catch (Exception ex)
        {
            Clients.showNotification(ex.getMessage());
        }

    }

    @Listen("onClick = #cancelSelectedProject")
    public void doCancelClick()
    {

        Messagebox.show("Está seguro que quiere volver?", "Confirmación",
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
