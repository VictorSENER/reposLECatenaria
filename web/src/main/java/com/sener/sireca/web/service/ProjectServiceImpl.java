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
    public int updateProject(Project project) throws Exception
    {

        // Check if title is empty.
        if (Strings.isBlank(project.getTitulo()))
            throw new Exception("El título del proyecto no puede estar vacío.");

        // Check if title is too long.
        else if (project.getTitulo().length() > 100)
            throw new Exception("El título del proyecto no puede ser tan largo. (Máximo 100 carácteres)");

        else if (getProjectByTitle(project.getTitulo()) != null
                && getProjectByTitle(project.getTitulo()).getId() != project.getId())
            throw new Exception("El título del proyecto ya existe.");

        // Check if client name is empty.
        if (Strings.isBlank(project.getCliente()))
            throw new Exception("El nombre del cliente no puede estar vacío.");

        // Check if client name is too long.
        else if (project.getCliente().length() > 50)
            throw new Exception("El nombre del cliente no puede ser tan largo. (Máximo 50 carácteres)");

        // Check if reference is empty.
        if (Strings.isBlank(project.getReferencia()))
            throw new Exception("La referencia no puede estar vacía.");

        // Check if reference is too long.
        else if (project.getReferencia().length() > 20)
            throw new Exception("La referencia no puede ser tan larga. (Máximo 20 carácteres)");

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
