package com.th.trust.data.util;

import java.util.function.Function;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.test.context.junit4.SpringRunner;

import com.th.trust.data.beijing.pojo.entity.BeijingCategoryIVitroDiagnosticReagentsRecord;

@RunWith(SpringRunner.class)
@SpringBootTest
public class MongoPageTest {

    @Autowired
    private MongoPageHelper mongoPageHelper;

    @Test
    public void aaaa() {

        final Query query = new Query(Criteria.where("jobId").is("0e1796ad-e0e5-4e02-9dfd-dbba701b5814"));

        PageResult<BeijingCategoryIVitroDiagnosticReagentsRecord> pageQuery = mongoPageHelper.pageQuery(query,
                BeijingCategoryIVitroDiagnosticReagentsRecord.class, 20, 1);
        PageResult<BeijingCategoryIVitroDiagnosticReagentsRecord> pageQuery2 = mongoPageHelper.pageQuery(query,
                BeijingCategoryIVitroDiagnosticReagentsRecord.class,Function.identity(), 20, 1,"");

        System.out.println(pageQuery);
        System.out.println(pageQuery.getList());
    }
}
