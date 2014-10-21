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

        // Checks if username is empty.
        if (Strings.isBlank(user.getUsername()))
        {
            Clients.showNotification("El nombre de usuario no puede estar vacío.");
            return 0;
        }

        else if (user.getUsername().length() > 50)
        {
            Clients.showNotification("El nombre de usuario no puede ser tan largo. (Máximo 50 carácteres)");
            return 0;
        }
        else if (getUserByUsername(user.getUsername()) != null)
        {
            Clients.showNotification("El nombre de usuario ya existe.");
            return 0;
        }

        // Checks if password is empty.
        if (Strings.isBlank(user.getPassword()))
        {
            Clients.showNotification("El password no puede estar vacío.");
            return 0;
        }

        else if (user.getPassword().length() > 50)
        {
            Clients.showNotification("El password no puede ser tan largo. (Máximo 50 carácteres)");
            return 0;
        }

        return userDao.updateUser(user);
    }

    @Override
    public int deleteUser(int id)
    {
        return userDao.deleteUser(id);
    }

}
