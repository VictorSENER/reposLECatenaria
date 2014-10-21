/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.dao;

import java.util.List;

import com.sener.sireca.web.bean.User;

public interface UserDao
{
    public int insertUser(User user);

    public List<User> getAllUsers();

    public User getUserById(int id);

    public int updateUser(User user);

    public int deleteUser(int id);
}