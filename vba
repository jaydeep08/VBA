import React from "react";

const Card = () => {
  const cardStyle = {
    backgroundImage: `url('https://via.placeholder.com/300')`, // Replace with your image URL
    backgroundSize: "cover",
    backgroundPosition: "center",
    borderRadius: "10px",
    padding: "20px",
    width: "300px",
    height: "200px",
    color: "white",
    display: "flex",
    flexDirection: "column",
    justifyContent: "flex-start",
    alignItems: "flex-start",
    boxShadow: "0 4px 6px rgba(0, 0, 0, 0.1)",
  };

  const buttonStyle = {
    backgroundColor: "blue",
    color: "white",
    border: "none",
    borderRadius: "5px",
    padding: "10px",
    width: "100%", // Take full width of the card
    marginTop: "10px",
    cursor: "pointer",
    fontSize: "16px",
  };

  return (
    <div style={cardStyle}>
      <h2>Card Heading</h2>
      <div style={{ marginTop: "auto", width: "100%" }}>
        <button style={buttonStyle}>Button 1</button>
        <button style={buttonStyle}>Button 2</button>
        <button style={buttonStyle}>Button 3</button>
      </div>
    </div>
  );
};

export default Card;
