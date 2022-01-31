import * as azure from "@pulumi/azure";
import * as storage from "azure-storage";
import { v4 } from "uuid";

// Create an Azure Resource Group
const resourceGroup = new azure.core.ResourceGroup("resourceGroup", {
    name: "to-do",
    location: "WestEurope",
  });

// Create an Azure resource (Storage Account)
const account = new azure.storage.Account("storage", {
    // The location for the storage account will be derived automatically from the resource group.
    resourceGroupName: resourceGroup.name,
    accountTier: "Standard",
    accountReplicationType: "LRS",
  });

// Export the connection string for the storage account
export const connectionString = account.primaryConnectionString;

const tasksTable = new azure.storage.Table("taskstable", {
  storageAccountName: account.name,
  name: "tasks",
});

//FUNCTIONS
//save task function
const addTaskFunction = new azure.appservice.HttpFunction("add", {
  route: "tasks",
  methods: ["POST"],
  callback: async (context, req) => {
    const body = req.body;
    const connString = connectionString.get();
    const tasksTableName = tasksTable.name.get();
    const requestId = v4();

    // variables from input
    const taskName = body.name;
    const taskCreated = body.created;
    const taskIsDone = body.isDone;

    // save to database
    const tableService = storage.createTableService(connString);
    const entGen = storage.TableUtilities.entityGenerator;
    const entry = {
      PartitionKey: entGen.String(requestId),
      RowKey: entGen.String("1"),
      name: entGen.String(taskName),
      created: entGen.String(taskCreated),
      isDone: entGen.Boolean(taskIsDone),
    };
    await new Promise((resolve, reject) => {
      tableService.insertEntity(
        tasksTableName,
        entry,
        function (error, result, response) {
          if (!error) {
            resolve(result);
          } else {
            reject(error);
          }
        }
      );
    });

    // output
    return {
      status: 201,
      body: {
        requestId: requestId,
      },
    };
  },
});

const getTasksFunction = new azure.appservice.HttpFunction("tasks", {
  route: "tasks",
  methods: ["GET"],
  callback: async (context, req) => {
    const connString = connectionString.get();
    const tasksTableName = tasksTable.name.get();    

    const tableService = storage.createTableService(connString);
    const query = new storage.TableQuery();

    const result = await new Promise((resolve, reject) => {
      tableService.queryEntities(
        tasksTableName,
        query,
        (null as unknown) as storage.TableService.TableContinuationToken,
        function (error, result, response) {
          if (error) {
            reject(error);
          } else {
            resolve(result);
          }
        }
      );
    });
    console.log((result as any).entries);
    const entries = (result as any).entries.map((x: any) => ({
      _id: x.PartitionKey["_"],
      name: x.name["_"],
      created: x.created["_"],
      isDone: x.isDone["_"],
    }));

    return {
      status: 200,
      body: entries,
    };
  },
});

const updateTaskFunction = new azure.appservice.HttpFunction("update", {
  route: "tasks",
  methods: ["PUT"],
  callback: async (context, req) => {
    const connString = connectionString.get();
    const tasksTableName = tasksTable.name.get();
    const requestId = req.body._id;

    // update database
    const tableService = storage.createTableService(connString);
    const entGen = storage.TableUtilities.entityGenerator;
    const entry = {
      PartitionKey: entGen.String(requestId),
      RowKey: entGen.String("1"),
      isDone: entGen.Boolean(true),
    };
    await new Promise((resolve, reject) => {
      tableService.mergeEntity(
        tasksTableName,
        entry,
        function (error, result, response) {
          if (!error) {
            resolve(result);
          } else {
            reject(error);
          }
        }
      );
    });

    // output
    return {
      status: 200,
      body: {
        message: "updated",
      },
    };
  },
});

const deleteTaskFunction = new azure.appservice.HttpFunction("delete", {
  route: "tasks/{id}",
  methods: ["DELETE"],
  callback: async (context, req) => {
    const connString = connectionString.get();
    const tasksTableName = tasksTable.name.get();
    const requestId = req.params.id;

    // delete from database
    const tableService = storage.createTableService(connString);
    const entGen = storage.TableUtilities.entityGenerator;
    const entry = {
      PartitionKey: entGen.String(requestId),
      RowKey: entGen.String("1")
    };
    await new Promise((resolve, reject) => {
      tableService.deleteEntity(
        tasksTableName,
        entry,
        function (error, response) {
          if (!error) {
            resolve(response);
          } else {
            reject(error);
          }
        }
      );
    });

    // output
    return {
      status: 204,
      body: {
        message: "deleted",
      },
    };
  },
});

new azure.appservice.MultiCallbackFunctionApp("application", {
  resourceGroupName: resourceGroup.name,
  functions: [addTaskFunction, getTasksFunction, updateTaskFunction, deleteTaskFunction],
  siteConfig:{
    cors:{
      allowedOrigins: ["*"],
    }
  }
});