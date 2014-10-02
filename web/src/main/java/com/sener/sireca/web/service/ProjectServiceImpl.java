/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Service;

import com.sener.sireca.web.bean.Globals;
import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.dao.ProjectDao;
import com.sener.sireca.web.util.SpringApplicationContext;

@Service("projectService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class ProjectServiceImpl implements ProjectService

{

    @Autowired
    ProjectDao projectDao;

    @Override
    public int insertProject(Project project)
    {

        int id = projectDao.insertProject(project);
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
        String ruta = System.getenv("SIRECA_HOME") + "/projects/" + id;

        // Crear carpetas
        fileService.addDirectory(ruta + Globals.CALCULO_REPLANTEO + "/1");
        fileService.addDirectory(ruta + Globals.DIBUJO_REPLANTEO + "/1");
        fileService.addDirectory(ruta + Globals.FICHAS_MONTAJE + "/1");
        fileService.addDirectory(ruta + Globals.FICHAS_PENDOLADO + "/1");

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
        return projectDao.updateProject(project);
    }

    @Override
    public int deleteProject(int id)
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
        fileService.deleteDirectory(System.getenv("SIRECA_HOME") + "/projects/"
                + id);

        return projectDao.deleteProject(id);
    }

    public void addDirectory(String ruta)
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
        fileService.addDirectory(ruta);

    }

    public boolean deleteDirectory(String ruta)
    {
        FileService fileService = (FileService) SpringApplicationContext.getBean("fileService");
        return fileService.deleteDirectory(ruta);
    }

}
