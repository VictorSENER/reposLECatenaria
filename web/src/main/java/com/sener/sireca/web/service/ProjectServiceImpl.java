/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Service;
import org.zkoss.lang.Strings;
import org.zkoss.zk.ui.util.Clients;

import com.sener.sireca.web.bean.DibujoVersion;
import com.sener.sireca.web.bean.MontajeVersion;
import com.sener.sireca.web.bean.PendoladoVersion;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.bean.ReplanteoVersion;
import com.sener.sireca.web.dao.ProjectDao;
import com.sener.sireca.web.util.IsJUnit;
import com.sener.sireca.web.util.SpringApplicationContext;

@Service("projectService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class ProjectServiceImpl implements ProjectService

{
    @Autowired
    ProjectDao projectDao;

    FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");

    @Override
    public int insertProject(Project project)
    {
        int id = projectDao.insertProject(project);
        String ruta = System.getenv("SIRECA_HOME");

        if (!IsJUnit.isJunitRunning())
            ruta += "/projects/";
        else
            ruta += "/projectTest/";

        ruta += id;

        // Crear carpetas
        fileService.addDirectory(ruta + ReplanteoVersion.CALCULO_REPLANTEO
                + "/1");
        fileService.addDirectory(ruta + DibujoVersion.DIBUJO_REPLANTEO + "/1");
        fileService.addDirectory(ruta + MontajeVersion.FICHAS_MONTAJE + "/1");
        fileService.addDirectory(ruta + PendoladoVersion.FICHAS_PENDOLADO
                + "/1");

        return id;
    }

    @Override
    public List<Project> getAllProjects()
    {
        return projectDao.getAllProjects();
    }

    @Override
    public Project getProjectById(int id)
    {
        for (Project p : getAllProjects())
            if (p.getId() == id)
                return p;

        return null;
    }

    @Override
    public Project getProjectByTitle(String title)
    {
        for (Project p : getAllProjects())
            if (p.getTitulo().equals(title))
                return p;

        return null;
    }

    @Override
    public int updateProject(Project project)
    {

        // Check if title is empty.
        if (Strings.isBlank(project.getTitulo()))
        {
            Clients.showNotification("El título del proyecto no puede estar vacío.");
            return 0;
        }
        // Check if title is too long.
        else if (project.getTitulo().length() > 100)
        {
            Clients.showNotification("El título del proyecto no puede ser tan largo. (Máximo 100 carácteres)");
            return 0;
        }
        else if (getProjectByTitle(project.getTitulo()) != null)
        {
            Clients.showNotification("El título del proyecto ya existe.");
            return 0;
        }

        // Check if client name is empty.
        if (Strings.isBlank(project.getCliente()))
        {
            Clients.showNotification("El nombre del cliente no puede estar vacío.");
            return 0;
        }

        // Check if client name is too long.
        else if (project.getCliente().length() > 50)
        {
            Clients.showNotification("El nombre del cliente no puede ser tan largo. (Máximo 50 carácteres)");
            return 0;
        }

        // Check if reference is empty.
        if (Strings.isBlank(project.getReferencia()))
        {
            Clients.showNotification("La referencia no puede estar vacía.");
            return 0;
        }
        // Check if reference is too long.
        else if (project.getReferencia().length() > 20)
        {
            Clients.showNotification("La referencia no puede ser tan larga. (Máximo 20 carácteres)");
            return 0;
        }

        return projectDao.updateProject(project);
    }

    @Override
    public int deleteProject(int id)
    {
        String ruta = System.getenv("SIRECA_HOME");

        if (!IsJUnit.isJunitRunning())
            ruta += "/projects/";
        else
            ruta += "/projectTest/";

        ruta += id;

        fileService.deleteDirectory(ruta);

        return projectDao.deleteProject(id);
    }

}
