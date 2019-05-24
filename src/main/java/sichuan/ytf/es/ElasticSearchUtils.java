package sichuan.ytf.es;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import org.apache.commons.lang3.StringUtils;
import org.elasticsearch.ElasticsearchException;
import org.elasticsearch.action.DocWriteResponse;
import org.elasticsearch.action.admin.indices.close.CloseIndexRequest;
import org.elasticsearch.action.admin.indices.create.CreateIndexRequest;
import org.elasticsearch.action.admin.indices.create.CreateIndexResponse;
import org.elasticsearch.action.admin.indices.delete.DeleteIndexRequest;
import org.elasticsearch.action.admin.indices.get.GetIndexRequest;
import org.elasticsearch.action.admin.indices.open.OpenIndexRequest;
import org.elasticsearch.action.admin.indices.open.OpenIndexResponse;
import org.elasticsearch.action.bulk.BulkItemResponse;
import org.elasticsearch.action.bulk.BulkRequest;
import org.elasticsearch.action.bulk.BulkResponse;
import org.elasticsearch.action.delete.DeleteRequest;
import org.elasticsearch.action.delete.DeleteResponse;
import org.elasticsearch.action.index.IndexRequest;
import org.elasticsearch.action.index.IndexResponse;
import org.elasticsearch.action.search.SearchRequest;
import org.elasticsearch.action.search.SearchResponse;
import org.elasticsearch.action.support.WriteRequest;
import org.elasticsearch.action.support.master.AcknowledgedResponse;
import org.elasticsearch.action.support.replication.ReplicationResponse;
import org.elasticsearch.action.update.UpdateRequest;
import org.elasticsearch.client.RequestOptions;
import org.elasticsearch.client.RestHighLevelClient;
import org.elasticsearch.common.settings.Settings;
import org.elasticsearch.common.unit.TimeValue;
import org.elasticsearch.common.xcontent.XContentType;
import org.elasticsearch.index.query.QueryBuilder;
import org.elasticsearch.rest.RestStatus;
import org.elasticsearch.search.builder.SearchSourceBuilder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import sichuan.ytf.web.BusinessException;

@Component
@SuppressWarnings("deprecation")
public class ElasticSearchUtils {
    private Logger logger = LoggerFactory.getLogger(ElasticSearchUtils.class);

    public static final String INDEX_KEY = "index";
    public static final String TYPE_KEY = "type";
    public static final String ID = "id";
    public static final String DATA = "data";

    @Autowired
    private RestHighLevelClient client;

    /**
     * 搜索
     * 
     * @param query {@link QueryBuilder}
     * @param index 索引名称
     * @return
     * @throws BusinessException
     */
    public List<Map<String, Object>> search(QueryBuilder query, String... index) throws BusinessException {
        SearchRequest searchRequest = new SearchRequest(index);
        SearchSourceBuilder sourceBuilder = new SearchSourceBuilder();
        sourceBuilder.query(query);
        sourceBuilder.timeout(new TimeValue(60, TimeUnit.SECONDS));
        // 前100页
        sourceBuilder.from(0);
        sourceBuilder.size(100);
        searchRequest.source(sourceBuilder);
        try {
            SearchResponse response = client.search(searchRequest, RequestOptions.DEFAULT);
            List<Map<String, Object>> list = new ArrayList<>();
            Arrays.stream(response.getHits().getHits()).forEach(i -> {
                Map<String, Object> source = i.getSourceAsMap();
                source.put("_score", i.getScore());
                source.put("_index", i.getIndex());
                source.put("_type", i.getType());
                source.put("_id", i.getId());
                list.add(source);
            });
            System.out.println("搜索结果总命中数：" + response.getHits().getTotalHits());
            return list;
        } catch (IOException e) {
            logger.error("查询elasticsearch失败，", e);
            throw new BusinessException("100", "查询elasticsearch失败，" + e.getMessage());
        }
    }

    /**
     * 验证索引是否存在
     * 
     * @param index 索引名称
     * @return
     * @throws IOException
     * @throws Exception
     */
    public boolean existIndex(String index) throws IOException {
        GetIndexRequest request = new GetIndexRequest();
        request.indices(index);
        request.local(false);
        request.humanReadable(true);
        boolean exists = client.indices().exists(request, RequestOptions.DEFAULT);
        return exists;
    }

    /**
     * 删除索引
     * 
     * @param index
     * @return
     * @throws Exception
     */
    public boolean deleteIndex(String... index) throws Exception {
        try {
            DeleteIndexRequest request = new DeleteIndexRequest(index);
            AcknowledgedResponse deleteIndexResponse = client.indices().delete(request, RequestOptions.DEFAULT);
            return deleteIndexResponse.isAcknowledged();
        } catch (ElasticsearchException exception) {
            if (exception.status() == RestStatus.NOT_FOUND) {
                logger.info("索引不存在");
                return true;
            } else {
                return false;
            }
        }
    }

    /**
     * 开启索引
     * 
     * @param indexName
     * @throws Exception
     */
    public void openIndex(String indexName) throws IOException {
        if (!existIndex(indexName)) {
            logger.error("索引不存在！");
            return;
        }
        OpenIndexRequest request = new OpenIndexRequest(indexName);
        OpenIndexResponse response = client.indices().open(request, RequestOptions.DEFAULT);
        if (response.isAcknowledged() || response.isShardsAcknowledged()) {
            logger.info("{} 索引开启成功！", indexName);
        }
    }

    /**
     * 关闭索引
     * 
     * @param indexName
     * @throws IOException
     */
    public void closeIndex(String indexName) throws IOException {
        if (!existIndex(indexName)) {
            logger.error("索引不存在！");
            return;
        }
        CloseIndexRequest request = new CloseIndexRequest(indexName);
        AcknowledgedResponse response = client.indices().close(request, RequestOptions.DEFAULT);
        if (response.isAcknowledged()) {
            logger.info("{} 索引已关闭！", indexName);
        }
    }

    /**
     * 
     * @param index
     * @param indexType
     * @param properties 结构: {name:{type:text}} {age:{type:integer}}
     * @return
     * @throws Exception
     */
    public boolean indexCreate(String index, String indexType, String propertyJson)
            throws BusinessException, Exception {

        if (existIndex(index)) {
            throw new BusinessException("100", "索引已存在");
        }
        CreateIndexRequest request = new CreateIndexRequest(index);
        // 为索引设置一个别名
//        request.alias(new Alias(index + indexType));
        // 创建的每个索引都可以有与之关联的特定设置。
        request.settings(Settings.builder().put("index.number_of_shards", 3).put("index.number_of_replicas", 2));
        request.mapping(indexType, propertyJson, XContentType.JSON);

        CreateIndexResponse createIndexResponse = client.indices().create(request, RequestOptions.DEFAULT);
        boolean acknowledged = createIndexResponse.isAcknowledged();
        return acknowledged;

    }

    /**
     * 
     * @param index
     * @param indexType
     * @param properties 结构: {name:{type:text}} {age:{type:integer}}
     * @return
     * @throws Exception
     */
    public boolean indexCreate(String index, String indexType, Map<String, Object> properties)
            throws BusinessException, Exception {

        if (existIndex(index)) {
            throw new BusinessException("100", "索引已存在");
        }
        CreateIndexRequest request = new CreateIndexRequest(index);
        request.settings(Settings.builder().put("index.number_of_shards", 3).put("index.number_of_replicas", 2));

        Map<String, Object> jsonMap = new HashMap<>();
        Map<String, Object> mapping = new HashMap<>();
        mapping.put("properties", properties);
        jsonMap.put(indexType, mapping);
        request.mapping(indexType, jsonMap);

        CreateIndexResponse createIndexResponse = client.indices().create(request, RequestOptions.DEFAULT);
        boolean acknowledged = createIndexResponse.isAcknowledged();
        return acknowledged;
    }

    /**
     * 创建更新文档，一条记录
     * 
     * @param index
     * @param indexType
     * @param documentId
     * @param map
     * @return
     * @throws Exception
     */
    public boolean addDocumentByMap(String index, String indexType, String documentId, Map<String, Object> map)
            throws Exception {
        IndexRequest request = new IndexRequest(index, indexType, documentId);

        request.source(map);
        IndexResponse indexResponse = client.index(request, RequestOptions.DEFAULT);

        if (indexResponse.getResult() == DocWriteResponse.Result.CREATED
                || indexResponse.getResult() == DocWriteResponse.Result.UPDATED) {
            return true;
        }
        ReplicationResponse.ShardInfo shardInfo = indexResponse.getShardInfo();
        if (shardInfo.getTotal() != shardInfo.getSuccessful()) {
            return true;
        }
        if (shardInfo.getFailed() > 0) {
            for (ReplicationResponse.ShardInfo.Failure failure : shardInfo.getFailures()) {
                throw new Exception(failure.reason());
            }
        }
        return false;
    }

    /**
     * 创建更新文档，一条记录
     * 
     * @param index
     * @param indexType
     * @param documentId
     * @param josonStr
     * @return
     * @throws Exception
     */
    public boolean addDocumentByJson(String index, String indexType, String documentId, String josonStr)
            throws Exception {
        IndexRequest request = new IndexRequest(index, indexType, documentId);

        request.source(josonStr, XContentType.JSON);
        IndexResponse indexResponse = client.index(request, RequestOptions.DEFAULT);

        if (indexResponse.getResult() == DocWriteResponse.Result.CREATED
                || indexResponse.getResult() == DocWriteResponse.Result.UPDATED) {
            return true;
        }
        ReplicationResponse.ShardInfo shardInfo = indexResponse.getShardInfo();
        if (shardInfo.getTotal() != shardInfo.getSuccessful()) {
            return true;
        }
        if (shardInfo.getFailed() > 0) {
            for (ReplicationResponse.ShardInfo.Failure failure : shardInfo.getFailures()) {
                throw new Exception(failure.reason());
            }
        }
        return false;
    }

    /**
     * 创建文档
     * 
     * @param index
     * @param indexType
     * @param jsonStr
     * @return
     * @throws Exception
     */
    public String documentCreate(String index, String indexType, String jsonStr) throws Exception {
        IndexRequest request = new IndexRequest(index, indexType);

        request.source(jsonStr, XContentType.JSON);
        IndexResponse indexResponse = client.index(request, RequestOptions.DEFAULT);

        String id = indexResponse.getId();
        if (indexResponse.getResult() == DocWriteResponse.Result.CREATED
                || indexResponse.getResult() == DocWriteResponse.Result.UPDATED) {
            return id;
        }
        ReplicationResponse.ShardInfo shardInfo = indexResponse.getShardInfo();
        if (shardInfo.getTotal() != shardInfo.getSuccessful()) {
            return id;
        }
        if (shardInfo.getFailed() > 0) {
            for (ReplicationResponse.ShardInfo.Failure failure : shardInfo.getFailures()) {
                throw new Exception(failure.reason());
            }
        }
        return null;
    }

    /**
     * 创建文档
     * 
     * @param index
     * @param indexType
     * @param map
     * @return
     * @throws Exception
     */
    public String documentCreate(String index, String indexType, Map<String, Object> map) throws Exception {
        IndexRequest request = new IndexRequest(index, indexType);

        request.source(map);
        IndexResponse indexResponse = client.index(request, RequestOptions.DEFAULT);

        String id = indexResponse.getId();
        if (indexResponse.getResult() == DocWriteResponse.Result.CREATED
                || indexResponse.getResult() == DocWriteResponse.Result.UPDATED) {
            return id;
        }
        ReplicationResponse.ShardInfo shardInfo = indexResponse.getShardInfo();
        if (shardInfo.getTotal() != shardInfo.getSuccessful()) {
            return id;
        }
        if (shardInfo.getFailed() > 0) {
            for (ReplicationResponse.ShardInfo.Failure failure : shardInfo.getFailures()) {
                throw new Exception(failure.reason());
            }
        }
        return null;
    }

    /**
     * 删除某条记录
     * 
     * @param index
     * @param indexType
     * @param documentId
     * @return
     * @throws Exception
     */
    public boolean documentDelete(String index, String indexType, String documentId) throws Exception {
        DeleteRequest request = new DeleteRequest(index, indexType, documentId);
        DeleteResponse deleteResponse = client.delete(request, RequestOptions.DEFAULT);
        if (deleteResponse.getResult() == DocWriteResponse.Result.NOT_FOUND) {
            return true;
        }
        ReplicationResponse.ShardInfo shardInfo = deleteResponse.getShardInfo();
        if (shardInfo.getTotal() != shardInfo.getSuccessful()) {
            return true;
        }
        if (shardInfo.getFailed() > 0) {
            for (ReplicationResponse.ShardInfo.Failure failure : shardInfo.getFailures()) {
                throw new Exception(failure.reason());
            }
        }
        return false;
    }

    /**
     * 批量增加文档 Map<String, String> 一条doc记录
     * index:indexName,type:typeName,id:docmentId,data:docmentJson
     * 
     * @param params
     * @throws IOException
     */
    public void bulkAdd(List<Map<String, String>> params) throws IOException {
        BulkRequest bulkRequest = new BulkRequest();
        for (Map<String, String> dataMap : params) {
            String index = dataMap.get(INDEX_KEY);
            String type = dataMap.get(TYPE_KEY);
            String id = dataMap.get(ID);
            String jsonString = dataMap.get(DATA);
            if (StringUtils.isNoneBlank(id, jsonString)) {
                IndexRequest request = new IndexRequest(index, type, id).source(jsonString, XContentType.JSON);
                bulkRequest.add(request);
            }
        }
        // 超时时间(2分钟)
        bulkRequest.timeout(TimeValue.timeValueMinutes(2L));
        // 刷新策略
        bulkRequest.setRefreshPolicy(WriteRequest.RefreshPolicy.WAIT_UNTIL);

        if (bulkRequest.numberOfActions() == 0) {
            logger.error("参数错误，批量增加操作失败！");
            return;
        }
        BulkResponse bulkResponse = client.bulk(bulkRequest, RequestOptions.DEFAULT);
        // 全部操作成功
        if (!bulkResponse.hasFailures()) {
            logger.info("批量增加操作成功！");
        } else {
            for (BulkItemResponse bulkItemResponse : bulkResponse) {
                if (bulkItemResponse.isFailed()) {
                    BulkItemResponse.Failure failure = bulkItemResponse.getFailure();
                    logger.error("\"index={}, type={}, id={}\"的文档增加失败！", failure.getIndex(), failure.getType(),
                            failure.getId());
                    logger.error("增加失败详情: {}", failure.getMessage());
                } else {
                    logger.info("\"index={}, type={}, id={}\"的文档增加成功！", bulkItemResponse.getIndex(),
                            bulkItemResponse.getType(), bulkItemResponse.getId());
                }
            }
        }
    }

    /**
     * 批量更新文档
     * 
     * @param params
     * @throws IOException
     */
    public void bulkUpdate(List<Map<String, String>> params) throws IOException {
        BulkRequest bulkRequest = new BulkRequest();
        for (Map<String, String> dataMap : params) {
            String index = dataMap.get(INDEX_KEY);
            String type = dataMap.get(TYPE_KEY);
            String id = dataMap.get(ID);
            String jsonString = dataMap.get(DATA);
            if (StringUtils.isNoneBlank(id, jsonString)) {
                UpdateRequest request = new UpdateRequest(index, type, id).doc(jsonString, XContentType.JSON);
                request.docAsUpsert(true);
                bulkRequest.add(request);
            }
        }
        if (bulkRequest.numberOfActions() == 0) {
            logger.error("参数错误，批量更新操作失败！");
            return;
        }
        bulkRequest.timeout(TimeValue.timeValueMinutes(2L));
        bulkRequest.setRefreshPolicy(WriteRequest.RefreshPolicy.WAIT_UNTIL);
        BulkResponse bulkResponse = client.bulk(bulkRequest, RequestOptions.DEFAULT);
        if (!bulkResponse.hasFailures()) {
            logger.info("批量更新操作成功！");
        } else {
            for (BulkItemResponse bulkItemResponse : bulkResponse) {
                if (bulkItemResponse.isFailed()) {
                    BulkItemResponse.Failure failure = bulkItemResponse.getFailure();
                    logger.error("\"index={}, type={}, id={}\"的文档更新失败！", failure.getIndex(), failure.getType(),
                            failure.getId());
                    logger.error("更新失败详情: {}", failure.getMessage());
                } else {
                    logger.info("\"index={}, type={}, id={}\"的文档更新成功！", bulkItemResponse.getIndex(),
                            bulkItemResponse.getType(), bulkItemResponse.getId());
                }
            }
        }
    }
}