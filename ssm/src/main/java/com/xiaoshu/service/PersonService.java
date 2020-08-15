package com.xiaoshu.service;

import java.io.IOException;
import java.util.Date;
import java.util.List;

import javax.jms.Destination;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jms.core.JmsTemplate;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import com.alibaba.fastjson.JSONObject;
import com.github.pagehelper.PageHelper;
import com.github.pagehelper.PageInfo;
import com.github.pagehelper.util.StringUtil;
import com.xiaoshu.dao.CompanyMapper;
import com.xiaoshu.dao.PersonMapper;
import com.xiaoshu.dao.UserMapper;
import com.xiaoshu.entity.Company;

import com.xiaoshu.entity.Person;
import com.xiaoshu.entity.PersonVo;

import com.xiaoshu.entity.User;
import com.xiaoshu.entity.UserExample;
import com.xiaoshu.entity.UserExample.Criteria;

@Service
public class PersonService {

	@Autowired
	PersonMapper personMapper;
	@Autowired
	private JmsTemplate jmsTemplate;
	
	@Autowired
	private Destination queueTextDestination;
	
	
	public PageInfo<PersonVo>findPage(PersonVo personVo,Integer pageNum,Integer pageSize){
		PageHelper.startPage(pageNum, pageSize);
		List<PersonVo>list = personMapper.findList(personVo);
		return new PageInfo<>(list);
	}
	
	@Autowired
	CompanyMapper companyMapper;
	public List<Company>findCompany(){
		return companyMapper.selectAll();
	}
	
	public Person findName(String expressName){
		Person param = new Person();
		param.setExpressName(expressName);
		return personMapper.selectOne(param);
	}
	
	public void addPerson(Person person){
		person.setCreateTime(new Date());
		personMapper.insert(person);
		
		jmsTemplate.convertAndSend(queueTextDestination,JSONObject.toJSONString(person));
	}
	public void updatePerson(Person person){
		personMapper.updateByPrimaryKeySelective(person);
	}
	public void deletePerson(Integer id){
		personMapper.deleteByPrimaryKey(id);
	}
	
	public List<PersonVo> findList(PersonVo personVo) {
		// TODO Auto-generated method stub
		return personMapper.findList(personVo);
	}
	
	public void importPerson(MultipartFile personFile) throws InvalidFormatException, IOException{
		//导入业务
		//获取工作表
		Workbook workbook = WorkbookFactory.create(personFile.getInputStream());
		
		//工作对象
		Sheet sheet = workbook.getSheetAt(0);
		
		int rowNum = sheet.getLastRowNum();//获取总行数
		
		for (int i = 0; i < rowNum; i++) {
			//第一行是表头，过滤
			Row row = sheet.getRow(i+1);
			String expressName = row.getCell(0).toString();
			String sex = row.getCell(1).toString();
			String expressTrait = row.getCell(2).toString();
			Date entryTime = row.getCell(3).getDateCellValue();
			String cname = row.getCell(4).toString();
			
			if (sex.equals("男")&& "申通快递".equals(cname)) {
				//把数据封装实体类
				Person p = new Person();
				p.setExpressName(expressName);
				p.setSex(sex);;p.setExpressTrait(expressTrait);
				p.setEntryTime(new Date());
				
			
				Company param = new Company();
				param.setExpressName(cname);
				Company company =companyMapper.selectOne(param);
				p.setExpressTypeId(company.getId());
				personMapper.insert(p);
				
			}
			
			
	}
}
	
	public List<PersonVo>countPerson(){
		return personMapper.countPerson();
	}
	
	
	}
