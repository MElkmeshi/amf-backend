import express from "express";
import fileUpload from "express-fileupload";
import axios from "axios";
import fs from "fs/promises";
import cors from "cors";
// const requestpromise = require("request-promise");
import requestpromise from "request-promise";
import * as url from "url";
const __filename = url.fileURLToPath(import.meta.url);
const __dirname = url.fileURLToPath(new URL(".", import.meta.url));

const app = express();
const PORT = process.env.PORT || 3000;
const baseURL = "https://alwadq.ly";

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(fileUpload());
app.use("/images", express.static(__dirname + "/images"));

class AMF {
  static TenantName = "";
  static ClientID = "";
  static ClientSecret = "";
  static TenantID = "";
  static ApplicationID = "";
  static RefreshToken = "";
  static AccessToken;
  static IssueDate;
  SiteName: string;
  ListName: string;
  constructor(ListName: string, SiteName: string) {
    this.SiteName = SiteName;
    this.ListName = ListName;
  }
  static async generatetoken() {
    let result = await requestpromise({
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      form: {
        grant_type: "refresh_token",
        client_id: `${this.ClientID}@${this.TenantID}`,
        client_secret: this.ClientSecret,
        resource: `${this.ApplicationID}/${this.TenantName}.sharepoint.com@${this.TenantID}`,
        refresh_token: this.RefreshToken,
      },
      uri: `https://accounts.accesscontrol.windows.net/${this.TenantID}/tokens/OAuth/2`,
      method: "POST",
    });
    return JSON.parse(result).access_token;
  }
  static async getAccessToken() {
    if (!this.AccessToken || Date.now().valueOf() > this.IssueDate + 28500000) {
      console.log("Generating new token");
      this.IssueDate = Date.now().valueOf();
      this.AccessToken = await this.generatetoken();
    }
    return this.AccessToken;
  }
  async getListItems() {
    let result = await requestpromise({
      json: true,
      verbose: true,
      headers: {
        Authorization: `Bearer ${await AMF.getAccessToken()}`,
        "Content-Type": "application/json; odata=verbose",
        Accept: "application/json; odata=nometadata",
      },
      uri: encodeURI(
        `https://${AMF.TenantName}.sharepoint.com/sites/${this.SiteName}/_api/web/lists/GetByTitle('${this.ListName}')/items`
      ),
      method: "GET",
    });
    return result;
  }
  async addAttachment(id, buffer) {
    let result = await requestpromise({
      headers: {
        Authorization: `Bearer ${await AMF.getAccessToken()}`,
        "Content-Type": "application/json; odata=verbose",
        Accept: "application/json; odata=nometadata",
      },
      uri: encodeURI(
        `https://${AMF.TenantName}.sharepoint.com/sites/${this.SiteName}/_api/lists/GetByTitle('${this.ListName}')/items(${id})/AttachmentFiles/add(FileName='${id}.jpg')`
      ),
      body: buffer,
      method: "POST",
    });
  }
  async deleteAllAttachment(id) {
    const attachments = await requestpromise({
      json: true,
      verbose: true,
      headers: {
        Authorization: `Bearer ${await AMF.getAccessToken()}`,
        "Content-Type": "application/json; odata=verbose",
        Accept: "application/json; odata=nometadata",
      },
      uri: encodeURI(
        `https://${AMF.TenantName}.sharepoint.com/sites/${this.SiteName}/_api/web/lists/GetByTitle('${this.ListName}')/items(${id})/AttachmentFiles`
      ),
      method: "GET",
    });
    attachments.value.forEach(async (attachment) => {
      const fileName = attachment.FileName;
      const result = await requestpromise({
        json: true,
        verbose: true,
        headers: {
          Authorization: `Bearer ${await AMF.getAccessToken()}`,
          "Content-Type": "application/json; odata=verbose",
          Accept: "application/json; odata=nometadata",
          "X-HTTP-Method": "DELETE",
          "IF-MATCH": "*",
        },
        uri: encodeURI(
          `https://${AMF.TenantName}.sharepoint.com/sites/${this.SiteName}/_api/web/lists/GetByTitle('${this.ListName}')/items(${id})/AttachmentFiles('${fileName}')`
        ),
        method: "POST",
      });
    });
  }
}
let SiteName = "REDStores";
let ListName = "AMFProducts";
let products = new AMF(ListName, SiteName);

const addImageToFiles = async (ID) => {
  const id = parseInt(ID);
  let result = await requestpromise({
    json: true,
    verbose: true,
    headers: {
      Authorization: `Bearer ${await AMF.getAccessToken()}`,
      "Content-Type": "application/json; odata=verbose",
      Accept: "application/json; odata=nometadata",
    },
    uri: encodeURI(
      `https://${AMF.TenantName}.sharepoint.com/sites/${SiteName}/_api/web/lists/GetByTitle('${ListName}')/items(${id})/AttachmentFiles`
    ),
    method: "GET",
  });
  console.log(result);
  const fileName = result.value[0].FileName;
  result = await requestpromise({
    json: true,
    verbose: true,
    headers: {
      Authorization: `Bearer ${await AMF.getAccessToken()}`,
      "Content-Type": "application/json; odata=verbose",
      Accept: "application/json; odata=nometadata",
    },
    uri: encodeURI(
      `https://${AMF.TenantName}.sharepoint.com/sites/${SiteName}/_api/web/lists/GetByTitle('${ListName}')/items(${id})/AttachmentFiles('${fileName}')/$value`
    ),
    method: "GET",
    encoding: null,
  });
  fs.writeFile(`./images/${id}.png`, result);
};
const deleteImageFromFiles = async (id) => {
  fs.unlink(`./images/${id}.png`);
};
function manychatjson(json, cardnum = 10) {
  let num = json.length;
  //{if you want to edit} edit the num varibles to the length of the list you want to repart about
  const titlekey = "Title";
  const subtitlekey = "Brand";
  const image_urlkey = "pictureLink";
  const pricekey = "Price";
  let i;
  const manychat = {
    version: "v2",
    content: {
      messages: [],
      actions: [],
      quick_replies: [
        {
          type: "node",
          caption: "القائمة الرئيسية 🏡",
          target: "Home",
        },
        {
          type: "node",
          caption: "رجوع",
          target: "Back",
        },
        {
          type: "node",
          caption: "عرض السلة 🛒",
          target: "Cart",
        },
      ],
    },
  };
  for (i = 0; i < Math.floor(num / cardnum); i++) {
    //@ts-ignore
    manychat["content"]["messages"].push({
      type: "cards",
      elements: [],
      image_aspect_ratio: "horizontal",
    });
    for (let j = 0; j < cardnum; j++) {
      const title = json[i * cardnum + j]["Title"];
      const subtitle = json[i * cardnum + j][subtitlekey];
      const image_url = json[i * cardnum + j][image_urlkey];
      const price = json[i * cardnum + j][pricekey];
      //@ts-ignore
      manychat["content"]["messages"][i]["elements"].push({
        title: title + " السعر " + price + " دينار ",
        subtitle,
        image_url,
        action_url: "https://manychat.com",
        buttons: [],
      });
    }
  }
  if (num % cardnum > 0) {
    manychat["content"]["messages"].push({
      type: "cards",
      elements: [],
      image_aspect_ratio: "horizontal",
    });
    for (let j = 0; j < num % cardnum; j++) {
      const title = json[i * cardnum + j][titlekey];
      const subtitle = json[i * cardnum + j][subtitlekey];
      const image_url = json[i * cardnum + j][image_urlkey];
      const price = json[i * cardnum + j][pricekey];
      manychat["content"]["messages"][i]["elements"].push({
        title: title + " السعر " + price + " دينار ",
        subtitle,
        image_url,
        action_url: "https://manychat.com",
        buttons: [],
      });
    }
  }
  return manychat;
}

//sharepoint CRUD operations:

//get all products
app.get("/api/v1/amf", async (req, res) => {
  res.send(
    (await products.getListItems()).value.map((item) => ({
      Id: item.Id,
      Title: item.Title,
      Brand: item.OData__x0645__x0627__x0631__x0643__x06,
      Size: item.OData__x0627__x0644__x0645__x0642__x06,
      Color: item.OData__x0644__x0648__x0646_,
      Addtions: item.OData__x0627__x0636__x0627__x0641__x06,
      Price: item.OData__x0633__x0639__x0631_ || 0,
      pictureLink: baseURL + "/images/" + item.Id + ".png",
    }))
  );
});
//get one products
app.get("/api/v1/amf/:id", async (req, res) => {
  const id = parseInt(req.params.id);
  res.send(
    (await products.getListItems()).value
      .filter((item) => item.Id == id)
      .map((item) => ({
        Id: item.Id,
        Title: item.Title,
        Brand: item.OData__x0645__x0627__x0631__x0643__x06,
        Size: item.OData__x0627__x0644__x0645__x0642__x06,
        Color: item.OData__x0644__x0648__x0646_,
        Addtions: item.OData__x0627__x0636__x0627__x0641__x06,
        Price: item.OData__x0633__x0639__x0631_ || 0,
        pictureLink: baseURL + "/images/" + item.Id + ".png",
      }))[0]
  );
});
//get all products in a manychat gallery format
app.get("/api/v1/amf/bot", async (req, res) => {
  res.send(
    manychatjson(
      (await products.getListItems()).value.map((item) => ({
        Id: item.Id,
        Title: item.Title,
        Brand: item.OData__x0645__x0627__x0631__x0643__x06,
        Size: item.OData__x0627__x0644__x0645__x0642__x06,
        Color: item.OData__x0644__x0648__x0646_,
        Addtions: item.OData__x0627__x0636__x0627__x0641__x06,
        Price: item.OData__x0633__x0639__x0631_ || 0,
        pictureLink: baseURL + "/images/" + item.Id + ".png",
      }))
    )
  );
});
//get one product in a manychat gallery format
app.get("/api/v1/amf/:id/bot", async (req, res) => {
  const id = parseInt(req.params.id);
  res.send(
    manychatjson(
      (await products.getListItems()).value
        .filter((item) => item.Id == id)
        .map((item) => ({
          Id: item.Id,
          Title: item.Title,
          Brand: item.OData__x0645__x0627__x0631__x0643__x06,
          Size: item.OData__x0627__x0644__x0645__x0642__x06,
          Color: item.OData__x0644__x0648__x0646_,
          Addtions: item.OData__x0627__x0636__x0627__x0641__x06,
          Price: item.OData__x0633__x0639__x0631_ || 0,
          pictureLink: baseURL + "/images/" + item.Id + ".png",
        }))
    )
  );
});
//create new product
app.post("/api/v1/additem", async (req, res) => {
  const item = req.body;
  const newitem = {
    __metadata: { type: "SP.Data.AMF_x0020_ProductsListItem" },
    Title: item.Title,
    OData__x0645__x0627__x0631__x0643__x06: item.Brand,
    OData__x0627__x0644__x0645__x0642__x06: item.Size,
    OData__x0644__x0648__x0646_: item.Color,
    OData__x0627__x0636__x0627__x0641__x06: item.Addtions,
    OData__x0633__x0639__x0631_: item.Price,
  };
  let result = await axios.post(
    `https://${AMF.TenantName}.sharepoint.com/sites/${SiteName}/_api/web/lists/GetByTitle('${ListName}')/items`,
    newitem,
    {
      headers: {
        Authorization: `Bearer ${await AMF.getAccessToken()}`,
        "Content-Type": "application/json; odata=verbose",
        Accept: "application/json; odata=nometadata",
      },
    }
  );
});
//delete product
app.delete("/api/v1/amf/:id", async (req, res) => {
  const id = req.params.id;
  let result = await requestpromise({
    json: true,
    verbose: true,
    headers: {
      Authorization: `Bearer ${await AMF.getAccessToken()}`,
      "Content-Type": "application/json; odata=verbose",
      Accept: "application/json; odata=nometadata",
      "X-HTTP-Method": "DELETE",
      "IF-MATCH": "*",
    },
    uri: encodeURI(
      `https://${AMF.TenantName}.sharepoint.com/sites/${SiteName}/_api/web/lists/GetByTitle('${ListName}')/items(${id})`
    ),
    method: "POST",
  });
  deleteImageFromFiles(id);
  res.send("Deleted successfully The item with ID " + id);
});
//update product not the image
app.post("/api/v1/amf/:id", async (req, res) => {
  const id = req.params.id;
  const item = req.body;
  const newitem = {
    __metadata: { type: "SP.Data.AMF_x0020_ProductsListItem" },
    Title: item.Title,
    OData__x0645__x0627__x0631__x0643__x06: item.Brand,
    OData__x0627__x0644__x0645__x0642__x06: item.Size,
    OData__x0644__x0648__x0646_: item.Color,
    OData__x0627__x0636__x0627__x0641__x06: item.Addtions,
    OData__x0633__x0639__x0631_: item.Price,
  };
  let result = await requestpromise({
    json: true,
    verbose: true,
    headers: {
      Authorization: `Bearer ${await AMF.getAccessToken()}`,
      "Content-Type": "application/json; odata=verbose",
      Accept: "application/json; odata=nometadata",
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*",
    },
    uri: encodeURI(
      `https://${AMF.TenantName}.sharepoint.com/sites/${SiteName}/_api/web/lists/GetByTitle('${ListName}')/items(${id})`
    ),
    method: "POST",
    body: newitem,
  });
  res.send("done");
});
//update product image
app.post("/api/v1/amf/photo/:id", async (req, res) => {
  const id = parseInt(req.params.id);
  const image = req.files.Image.data;
  await products.deleteAllAttachment(id);
  deleteImageFromFiles(id);
  products.addAttachment(id, image);
  addImageToFiles(id);
  res.send("Photos updated successfully");
});
//add to postman
//update product with image
app.post("/api/v1/allamf/:id", async (req, res) => {
  const id = req.params.id;
  const item = req.body;
  const newitem = {
    __metadata: { type: "SP.Data.AMF_x0020_ProductsListItem" },
    Title: item.Title,
    OData__x0645__x0627__x0631__x0643__x06: item.Brand,
    OData__x0627__x0644__x0645__x0642__x06: item.Size,
    OData__x0644__x0648__x0646_: item.Color,
    OData__x0627__x0636__x0627__x0641__x06: item.Addtions,
    OData__x0633__x0639__x0631_: item.Price,
  };
  let result = await requestpromise({
    json: true,
    verbose: true,
    headers: {
      Authorization: `Bearer ${await AMF.getAccessToken()}`,
      "Content-Type": "application/json; odata=verbose",
      Accept: "application/json; odata=nometadata",
      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*",
    },
    uri: encodeURI(
      `https://${AMF.TenantName}.sharepoint.com/sites/${SiteName}/_api/web/lists/GetByTitle('${ListName}')/items(${id})`
    ),
    method: "POST",
    body: newitem,
  });
  const theID = parseInt(id);
  await products.deleteAllAttachment(theID);
  deleteImageFromFiles(theID);
  products.addAttachment(id, req.files.Image.data);
  addImageToFiles(theID);
  res.send("The new ID is " + result.ID);
});
//get product image from sharepoint
app.get("/api/v1/getphoto/:id", async (req, res) => {
  const id = parseInt(req.params.id);
  let result = await requestpromise({
    json: true,
    verbose: true,
    headers: {
      Authorization: `Bearer ${await AMF.getAccessToken()}`,
      "Content-Type": "application/json; odata=verbose",
      Accept: "application/json; odata=nometadata",
    },
    uri: encodeURI(
      `https://${AMF.TenantName}.sharepoint.com/sites/${SiteName}/_api/web/lists/GetByTitle('${ListName}')/items(${id})/AttachmentFiles`
    ),
    method: "GET",
  });
  const fileName = result.value[0].FileName;
  result = await requestpromise({
    json: true,
    verbose: true,
    headers: {
      Authorization: `Bearer ${await AMF.getAccessToken()}`,
      "Content-Type": "application/json; odata=verbose",
      Accept: "application/json; odata=nometadata",
    },
    uri: encodeURI(
      `https://${AMF.TenantName}.sharepoint.com/sites/${SiteName}/_api/web/lists/GetByTitle('${ListName}')/items(${id})/AttachmentFiles('${fileName}')/$value`
    ),
    method: "GET",
    encoding: null,
  });
  fs.writeFile(`./images/${id}.png`, result);
  res.contentType("image/jpeg");
  res.send(result);
});
//get all images from sharepoint and save them in the images folder
app.get("/allphotos", async (req, res) => {
  (await products.getListItems()).value.forEach(async (element) => {
    const id = element.Id;
    try {
      let result = await requestpromise({
        json: true,
        verbose: true,
        headers: {
          Authorization: `Bearer ${await AMF.getAccessToken()}`,
          "Content-Type": "application/json; odata=verbose",
          Accept: "application/json; odata=nometadata",
        },
        uri: encodeURI(
          `https://${AMF.TenantName}.sharepoint.com/sites/${SiteName}/_api/web/lists/GetByTitle('${ListName}')/items(${id})/AttachmentFiles`
        ),
        method: "GET",
      });
      const fileName = result.value[0].FileName;
      result = await requestpromise({
        json: true,
        verbose: true,
        headers: {
          Authorization: `Bearer ${await AMF.getAccessToken()}`,
          "Content-Type": "application/json; odata=verbose",
          Accept: "application/json; odata=nometadata",
        },
        uri: encodeURI(
          `https://${AMF.TenantName}.sharepoint.com/sites/${SiteName}/_api/web/lists/GetByTitle('${ListName}')/items(${id})/AttachmentFiles('${fileName}')/$value`
        ),
        method: "GET",
        encoding: null,
      });
      await fs.writeFile(`./images/${id}.png`, result);
    } catch (e) {
      console.log("error " + id);
      console.log(e);
    }
  });
  res.send("done");
});
app.listen(PORT, () => {
  console.log(`Example app listening at http://localhost:${PORT}`);
});
