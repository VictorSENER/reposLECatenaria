/**
 * Copyright(c) 2014 SENER Ingenieria y Sistemas SA All rights reserved.
 */

package com.sener.sireca.web.bean;

import javax.persistence.Column;
import javax.persistence.Entity;
import javax.persistence.GeneratedValue;
import javax.persistence.Id;
import javax.persistence.Table;

import org.hibernate.annotations.GenericGenerator;

import com.sener.sireca.web.util.IsJUnit;

@Entity
@Table(name = "Proyecto")
public class Project
{
    // Identificador del usuario
    @Id
    @GeneratedValue(generator = "increment")
    @GenericGenerator(name = "increment", strategy = "increment")
    private Integer id;

    // Titulo del proyecto
    @Column(name = "titulo")
    private String titulo;

    // Referencia del Proyecto
    @Column(name = "referencia")
    private String referencia;

    // Nombre del cliente
    @Column(name = "cliente")
    private String cliente;

    // Id del usuario
    @Column(name = "idUsuario")
    private int idUsuario;

    // Id de la catenaria
    @Column(name = "idCatenaria")
    private int idCatenaria;

    // Es v�a doble?
    @Column(name = "viaDoble")
    private boolean viaDoble;

    // Nombre plantilla de Pendolado
    @Column(name = "pendolado")
    private String pendolado;

    // Nombre plantilla de Montaje
    @Column(name = "montaje")
    private String montaje;

    public Integer getId()
    {
        return id;
    }

    public String getTitulo()
    {
        return titulo;
    }

    public void setTitulo(String titulo)
    {
        this.titulo = titulo;
    }

    public int getIdUsuario()
    {
        return idUsuario;
    }

    public void setIdUsuario(int idUsuario)
    {
        this.idUsuario = idUsuario;
    }

    public String getCliente()
    {
        return cliente;
    }

    public void setCliente(String cliente)
    {
        this.cliente = cliente;
    }

    public String getReferencia()
    {
        return referencia;
    }

    public void setReferencia(String referencia)
    {
        this.referencia = referencia;
    }

    public int getIdCatenaria()
    {
        return idCatenaria;
    }

    public void setIdCatenaria(int idCatenaria)
    {
        this.idCatenaria = idCatenaria;
    }

    public boolean getViaDoble()
    {
        return viaDoble;
    }

    public void setViaDoble(boolean viaDoble)
    {
        this.viaDoble = viaDoble;
    }

    public String getMontaje()
    {
        return montaje;
    }

    public void setMontaje(String montaje)
    {
        this.montaje = montaje;
    }

    public String getPendolado()
    {
        return pendolado;
    }

    public void setPendolado(String pendolado)
    {
        this.pendolado = pendolado;
    }

    public String getCalcReplanteoBasePath()
    {
        return getBasePath() + ReplanteoVersion.CALCULO_REPLANTEO;
    }

    public String getDibReplanteoBasePath()
    {
        return getBasePath() + DibujoVersion.DIBUJO_REPLANTEO;
    }

    public String getMonReplanteoBasePath()
    {
        return getBasePath() + MontajeVersion.FICHAS_MONTAJE;
    }

    public String getPenReplanteoBasePath()
    {
        return getBasePath() + PendoladoVersion.FICHAS_PENDOLADO;
    }

    private String getBasePath()
    {
        String basePath = System.getenv("SIRECA_HOME");

        if (!IsJUnit.isJunitRunning())
            basePath += "/projects/";
        else
            basePath += "/projectTest/";

        return basePath + id;
    }

    public String getTemplatePath(String fase)
    {

        return System.getenv("SIRECA_HOME") + "/templates" + fase;
    }

    public String getTemplate(String fase)
    {

        if (fase.equals("/fichas-pendolado/"))
            return getTemplatePath(fase) + pendolado;
        else
            return getTemplatePath(fase) + montaje;

    }

}
