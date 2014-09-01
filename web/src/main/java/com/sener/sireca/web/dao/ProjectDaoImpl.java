/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.dao;

import java.io.Serializable;
import java.util.List;

import org.hibernate.Session;
import org.hibernate.SessionFactory;
import org.hibernate.Transaction;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;

import com.sener.sireca.web.bean.Project;

@Repository("projectDao")
public class ProjectDaoImpl implements ProjectDao
{

    @Autowired
    SessionFactory sessionFactory;

    public int insertProject(Project project)
    {
        Session session = sessionFactory.openSession();
        Transaction tx = session.beginTransaction();
        session.saveOrUpdate(project);
        tx.commit();
        Serializable id = session.getIdentifier(project);
        session.close();
        return (Integer) id;
    }

    public List<Project> getAllProjects()
    {
        Session session = sessionFactory.openSession();
        @SuppressWarnings("unchecked")
        List<Project> projectList = session.createQuery("FROM Project").list();
        session.close();
        return projectList;
    }

    public Project getProjectById(int id)
    {
        Session session = sessionFactory.openSession();
        Project project = (Project) session.load(Project.class, id);
        session.close();
        return project;
    }

    public int updateProject(Project project)
    {
        Session session = sessionFactory.openSession();
        Transaction tx = session.beginTransaction();
        session.saveOrUpdate(project);
        tx.commit();
        Serializable id = session.getIdentifier(project);
        session.close();
        return (Integer) id;
    }

    public int deleteProject(int id)
    {
        Session session = sessionFactory.openSession();
        Transaction tx = session.beginTransaction();
        Project project = (Project) session.load(Project.class, id);
        session.delete(project);
        tx.commit();
        Serializable ids = session.getIdentifier(project);
        session.close();
        return (Integer) ids;
    }
}
