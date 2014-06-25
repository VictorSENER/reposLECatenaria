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

import com.sener.sireca.web.bean.User;

@Repository("userDao")
public class UserDaoImpl implements UserDao
{

    @Autowired
    SessionFactory sessionFactory;

    public int insertUser(User user)
    {
        Session session = sessionFactory.openSession();
        Transaction tx = session.beginTransaction();
        session.saveOrUpdate(user);
        tx.commit();
        Serializable id = session.getIdentifier(user);
        session.close();
        return (Integer) id;
    }

    public List<User> getAllUsers()
    {
        Session session = sessionFactory.openSession();
        @SuppressWarnings("unchecked")
        List<User> userList = session.createQuery("FROM User").list();
        session.close();
        return userList;
    }

    public User getUserById(int id)
    {
        Session session = sessionFactory.openSession();
        User user = (User) session.load(User.class, id);
        session.close();
        return user;
    }

    public int updateUser(User user)
    {
        Session session = sessionFactory.openSession();
        Transaction tx = session.beginTransaction();
        session.saveOrUpdate(user);
        tx.commit();
        Serializable id = session.getIdentifier(user);
        session.close();
        return (Integer) id;
    }

    public int deleteUser(int id)
    {
        Session session = sessionFactory.openSession();
        Transaction tx = session.beginTransaction();
        User user = (User) session.load(User.class, id);
        session.delete(user);
        tx.commit();
        Serializable ids = session.getIdentifier(user);
        session.close();
        return (Integer) ids;
    }
}
