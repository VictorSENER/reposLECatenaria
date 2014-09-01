/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.page;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.select.SelectorComposer;
import org.zkoss.zk.ui.select.annotation.Wire;
import org.zkoss.zul.Button;
import org.zkoss.zul.ListModelList;
import org.zkoss.zul.Textbox;

import com.sener.sireca.web.bean.Project;
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

    // Projects list
    ListModelList<Project> projectListModel;

    // Project currently selected.
    Project selectedProject;

    // Session data
    // HttpSession session = (HttpSession)
    // Sessions.getCurrent().getNativeSession();
    // AuthenticationService authService = (AuthenticationService)
    // SpringApplicationContext.getBean("authService");
    // ActiveProjectService actProj = (ActiveProjectService)
    // SpringApplicationContext.getBean("actProj");

    @Override
    public void doAfterCompose(Component comp) throws Exception
    {
        super.doAfterCompose(comp);

        // Fill projects list using DB data
        ProjectService projectService = (ProjectService) SpringApplicationContext.getBean("projectService");
        UserService userService = (UserService) SpringApplicationContext.getBean("userService");

        selectedProject = projectService.getProjectById(1);

        selectedProjectTitle.setText(selectedProject.getTitulo());
        selectedProjectClient.setText(selectedProject.getCliente());
        selectedProjectUser.setText(userService.getUserById(
                (selectedProject.getIdUsuario())).getUsername());
        selectedProjectReference.setText(selectedProject.getReferencia());

    }

}
