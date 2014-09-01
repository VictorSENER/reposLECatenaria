/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Service;

import com.sener.sireca.web.bean.Project;
import com.sener.sireca.web.dao.ProjectDao;

@Service("projectService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class ProjectServiceImpl implements ProjectService
{

    @Autowired
    ProjectDao projectDao;

    public int insertProject(Project project)
    {
        return projectDao.insertProject(project);
    }

    public List<Project> getAllProjects()
    {
        return projectDao.getAllProjects();
    }

    public Project getProjectById(int id)
    {
        return projectDao.getProjectById(id);
    }

    public Project getProjectByTitle(String title)
    {
        for (Project p : getAllProjects())
        {
            if (p.getTitulo().equals(title))
                return p;
        }

        return null;
    }

    public int updateProject(Project project)
    {
        return projectDao.updateProject(project);
    }

    public int deleteProject(int id)
    {
        return projectDao.deleteProject(id);
    }

}
