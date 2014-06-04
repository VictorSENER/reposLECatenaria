/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.sener.sireca.web.bean.User;
import com.sener.sireca.web.dao.UserDao;

@Service
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
        return userDao.getUserById(id);
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
