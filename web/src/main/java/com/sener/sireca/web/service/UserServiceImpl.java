/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Service;

import com.sener.sireca.web.bean.User;
import com.sener.sireca.web.dao.UserDao;

@Service("userService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class UserServiceImpl implements UserService
{

    @Autowired
    UserDao userDao;

    @Override
    public int insertUser(User user)
    {
        return userDao.insertUser(user);
    }

    @Override
    public List<User> getAllUsers()
    {
        return userDao.getAllUsers();
    }

    @Override
    public User getUserById(int id)
    {
        for (User u : getAllUsers())
            if (u.getId() == id)
                return u;

        return null;
    }

    @Override
    public User getUserByUsername(String username)
    {
        for (User u : getAllUsers())
            if (u.getUsername().equals(username))
                return u;

        return null;
    }

    @Override
    public int updateUser(User user)
    {
        return userDao.updateUser(user);
    }

    @Override
    public int deleteUser(int id)
    {
        return userDao.deleteUser(id);
    }

}
