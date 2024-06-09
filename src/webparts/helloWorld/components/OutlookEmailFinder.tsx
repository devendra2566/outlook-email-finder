import React, { useEffect, useState } from "react";
import { Web } from "sp-pnp-js";
import { FaInfoCircle } from "react-icons/fa";
import { FaFilter } from "react-icons/fa";

interface Person {
  Id: number;
  Title: string;
  Body: string;
  creationTime: string;
  recipients: string;
  senderEmail: string;
  Created: string;
  Author: { Title: string }
  Editor: { Title: string };
}

const OutlookEmailFinder = () => {
  const [data, setData] = useState<Person[]>([]);
  const [backupData, setBackupData] = useState<Person[]>([]);
  const [expanded, setExpanded] = useState<any>({});

  //here fetch the data froom the server
  useEffect(() => {
    const fetchEmailData = async () => {
      try {
        const getData = new Web(
          "https://smalsusinfolabs.sharepoint.com/sites/HHHHQA/SP"
        );
        const res = await getData.lists
          .getById("18c9128d-3710-4ceb-a714-9ce9d1a0dae4")
          .items.select(
            "Id",
            "Title",
            "Body",
            "creationTime",
            "Created",
            "Author/Title",
            "recipients",
            "Editor/Title",
            "senderEmail",
            "Portfolios/Title",
            "FileLeafRef",
            "FileDirRef",
            "File_x0020_Type"
          )
          .expand("Author", "Editor", "Portfolios")
          .getAll();

        const data1 = res?.filter(
          (item: any) => item.File_x0020_Type === "msg"
        );
        setData(data1);
        setBackupData(data1);
      } catch (error) {
        console.error(error, "error of fetch Data");
      }
    };

    fetchEmailData();
  }, []);

  const countWords = (str: string) => {
    return str?.split(" ").filter((word) => word.length > 0).length;
  };

  const emailColumnFilter = (data: string, columnName: string) => {
    if (columnName == "reciever") {
      const res = backupData.filter((item) => item.recipients === data);
      setData(res);
    } else {
      const res = backupData.filter((item) => item.senderEmail === data);
      setData(res);
    }
  };
  // const columnFilter = (sender:string) => {
  //   const res = backupData.filter((item) => item.senderEmail === sender);
  //   setData(res);
  // };
  const removeAllFilter = () => {
    setData(backupData);
  };

  const toggleInfo = (id: number) => {
    setExpanded({ ...expanded, [id]: !expanded[id] });
  };

  // handleSearch is used for serching
  const handleSearch = (e: any) => {
    const { name, value } = e.target;
    console.log(name, value);

    if (
      name == "Title" ||
      name == "senderEmail" ||
      name == "recipients" ||
      name == "Body" ||
      name == "creationTime" ||
      name == "created" ||
      name == "Editor"
    ) {
      const coulumnSearch = backupData.filter((item: any) =>
        item[name].toString().toLowerCase().includes(value.toLowerCase())
      );
      setData(coulumnSearch);
    } else {
      const globalSearch = backupData?.filter((item: any) =>
        Object?.values({ ...item })
          .toString()
          .toLowerCase()
          .includes(value.toLowerCase())
      );
      setData(globalSearch);
    }
  };

  //''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  return (
    <div

      onChange={handleSearch}
      style={{ border: "2px dashed black", borderRadius: "2px" }}
    > 
    <h1 style={{ marginLeft: "430px"}}>Outlook Email Finder</h1>
      <div style={{ cursor: "pointer", marginLeft: "1200px" }}>
        <FaFilter onClick={removeAllFilter} />
      </div>
      <input
        name="gloabal"
        placeholder="globalSearch.."
        style={{
          borderRadius: "2px",
          border: " 2px solid black",
          marginLeft: "500px",
        }}
      ></input>
      <table
        className="table table-striped"
        style={{ borderRadius: "10px solid black" }}
      >
        <thead>
          <tr>
            <td>
              {" "}
              <input
                name="Title"
                placeholder="Title"
                style={{
                  width: "140px",
                  height: "27px",
                  borderRadius: "2px",
                  border: " 2px solid black",
                }}
              ></input>
            </td>
            <td>
              <input
                name="senderEmail"
                placeholder="senderEmail"
                style={{
                  width: "140px",
                  height: "27px",
                  borderRadius: "2px",
                  border: " 2px solid black",
                }}
              ></input>
            </td>
            <td>
              {" "}
              <input
                name="recipients"
                placeholder="recipients"
                style={{
                  width: "140px",
                  height: "27px",
                  borderRadius: "2px",
                  border: " 2px solid black",
                }}
              ></input>
            </td>
            <td>
              {" "}
              <input
                name="Body"
                placeholder="Body"
                style={{
                  width: "120px",
                  height: "27px",
                  borderRadius: "2px",
                  border: " 2px solid black",
                }}
              ></input>
            </td>
            <td>
              {" "}
              <input
                name="creationTime"
                placeholder="creationTime"
                style={{
                  width: "140px",
                  height: "27px",
                  borderRadius: "2px",
                  border: " 2px solid black",
                }}
              ></input>
            </td>
            <td>
              {" "}
              <input
                name="created"
                placeholder="Author"
                style={{
                  width: "140px",
                  height: "27px",
                  borderRadius: "2px",
                  border: " 2px solid black",
                }}
              ></input>
            </td>
            <td>
              {" "}
              <input
                name="Editor"
                placeholder="Editor"
                style={{
                  width: "140px",
                  height: "27px",
                  borderRadius: "2px",
                  border: " 2px solid black",
                }}
              ></input>
            </td>
          </tr>
        </thead>
        <tbody>
          {data?.map((item: Person) => (
            <tr key={item?.Id}>
              <td>{item?.Title}</td>
              <td
                onClick={() => emailColumnFilter(item?.senderEmail, "sender")}
                style={{ cursor: "pointer" }}
              >
                {item?.senderEmail}
              </td>
              {/* <td onClick={() => columnFilter(item?.recipients,'reciever')} style={{cursor:"pointer"}}>{item.recipients}</td> */}
              <td
                onClick={() => emailColumnFilter(item?.recipients, "reciever")}
                style={{ cursor: "pointer" }}
              >
                {JSON.parse(item?.recipients)?.map((item: any) => (
                  <div>{item?.email}</div>
                ))}
              </td>
              {/* {countWords(item.Body) > 10 ? (
              <td>
                {item.Body.split(' ').slice(0, 10).join(' ')}
                <FaInfoCircle   />
              </td>
            ) : (
              <td>{item.Body}</td>
            )}  */}

              <td>
                {expanded[item?.Id]
                  ? item?.Body
                  : item?.Body?.split(" ").slice(0, 10).join(" ")}
                {countWords(item?.Body) > 10 && (
                  <FaInfoCircle
                    onClick={() => toggleInfo(item?.Id)}
                    style={{ cursor: "pointer", marginLeft: "5px" }}
                  />
                )}
              </td>
              {/* <td>{item?.Author?.Title}</td> */}

              <td>{item?.creationTime}</td>
              <td>
                {item?.Author?.Title ? item?.Author?.Title : item?.Created}
              </td>

              <td>{item?.Editor?.Title}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};

export default OutlookEmailFinder;
