/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.dao;

import java.util.List;

import org.hibernate.Session;
import org.hibernate.SessionFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Repository;

import com.sener.sireca.web.bean.Catenaria;

@Repository("catenariaDao")
public class CatenariaDaoImpl implements CatenariaDao
{

    @Autowired
    SessionFactory sessionFactory;

    public List<Catenaria> getAllCatenarias()
    {
        Session session = sessionFactory.openSession();
        @SuppressWarnings("unchecked")
        List<Catenaria> catenariaList = session.createQuery("FROM Catenaria").list();
        session.close();
        return catenariaList;
    }

    public Catenaria getCatenariaById(int id)
    {
        Session session = sessionFactory.openSession();
        Catenaria catenaria = (Catenaria) session.load(Catenaria.class, id);
        session.close();
        return catenaria;
    }

}
