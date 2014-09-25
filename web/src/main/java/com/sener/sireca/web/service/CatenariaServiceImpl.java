/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.service;

import java.util.ArrayList;
import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.context.annotation.ScopedProxyMode;
import org.springframework.stereotype.Service;

import com.sener.sireca.web.bean.Catenaria;
import com.sener.sireca.web.dao.CatenariaDao;

@Service("catenariaService")
@Scope(value = "singleton", proxyMode = ScopedProxyMode.TARGET_CLASS)
public class CatenariaServiceImpl implements CatenariaService

{

    @Autowired
    CatenariaDao catenariaDao;

    public List<Catenaria> getAllCatenarias()
    {
        return catenariaDao.getAllCatenarias();
    }

    public Catenaria getCatenariaById(int id)
    {
        for (Catenaria c : getAllCatenarias())
            if (c.getId() == id)
                return c;

        return null;
    }

    public Catenaria getCatenariaByTitle(String nombre)
    {
        for (Catenaria c : getAllCatenarias())
            if (c.getNomCatenaria().equals(nombre))
                return c;

        return null;
    }

    public List<String> getListCatenarias()
    {
        List<Catenaria> cat = catenariaDao.getAllCatenarias();

        ArrayList<String> catList = new ArrayList<String>();

        for (int i = 0; i < cat.size(); i++)
            catList.add(cat.get(i).getNomCatenaria());

        return catList;
    }
}
