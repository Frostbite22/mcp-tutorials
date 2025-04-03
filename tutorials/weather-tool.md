# Building a Weather Tool with MCP

This tutorial will walk you through building a basic weather service tool using the Model Context Protocol (MCP). This tool will allow AI models to fetch weather information for locations around the world.

## Prerequisites

- Node.js (v14+) installed
- Basic understanding of JavaScript and JSON-RPC
- OpenWeatherMap API key (free tier available at [openweathermap.org](https://openweathermap.org/))

## Project Setup

1. Create a new directory for your project:

```bash
mkdir mcp-weather-tool
cd mcp-weather-tool
```

2. Initialize a new Node.js project:

```bash
npm init -y
```

3. Install the required dependencies:

```bash
npm install express node-fetch
```

## Creating the MCP Server

Create a new file called `server.js` with the following content:

```javascript
const express = require('express');
const fetch = require('node-fetch');
const app = express();

// Replace with your own API key
const WEATHER_API_KEY = 'YOUR_OPENWEATHERMAP_API_KEY';
const WEATHER_API_BASE_URL = 'https://api.openweathermap.org/data/2.5/weather';

// Middleware to parse JSON requests
app.use(express.json());

// Handle JSON-RPC requests
app.post('/', async (req, res) => {
  const { jsonrpc, method, params, id } = req.body;
  
  // Verify JSON-RPC 2.0 request
  if (jsonrpc !== '2.0' || !method || !id) {
    return res.json({
      jsonrpc: '2.0',
      error: { code: -32600, message: 'Invalid request' },
      id: null
    });
  }
  
  // Handle method calls
  try {
    let result;
    
    switch (method) {
      case 'initialize':
        // Respond to initialization with capabilities
        result = {
          protocolVersion: '2024-11-05',
          methods: {
            getCurrentWeather: {
              description: 'Get current weather for a location',
              parameters: {
                type: 'object',
                properties: {
                  location: {
                    type: 'string',
                    description: 'City name, state code and country code divided by comma'
                  },
                  units: {
                    type: 'string',
                    enum: ['metric', 'imperial'],
                    description: 'Units of measurement',
                    default: 'metric'
                  }
                },
                required: ['location']
              }
            }
          }
        };
        break;
        
      case 'getCurrentWeather':
        // Check for required parameters
        if (!params || !params.location) {
          throw { code: -32602, message: 'Invalid params - location is required' };
        }
        
        // Set default units if not provided
        const units = params.units || 'metric';
        
        // Call the weather API
        const weatherUrl = `${WEATHER_API_BASE_URL}?q=${encodeURIComponent(params.location)}&units=${units}&appid=${WEATHER_API_KEY}`;
        const weatherResponse = await fetch(weatherUrl);
        const weatherData = await weatherResponse.json();
        
        if (weatherResponse.status !== 200) {
          throw { 
            code: -32603, 
            message: `Weather API error: ${weatherData.message || 'Unknown error'}` 
          };
        }
        
        // Format the weather data for our response
        result = {
          location: {
            name: weatherData.name,
            country: weatherData.sys.country,
            coordinates: {
              lat: weatherData.coord.lat,
              lon: weatherData.coord.lon
            }
          },
          current: {
            temperature: weatherData.main.temp,
            feels_like: weatherData.main.feels_like,
            humidity: weatherData.main.humidity,
            pressure: weatherData.main.pressure,
            wind: {
              speed: weatherData.wind.speed,
              degrees: weatherData.wind.deg
            },
            weather: {
              main: weatherData.weather[0].main,
              description: weatherData.weather[0].description,
              icon: weatherData.weather[0].icon
            }
          },
          units: units
        };
        break;
        
      default:
        // Handle unknown methods
        throw { code: -32601, message: `Method not found: ${method}` };
    }
    
    // Return successful response
    return res.json({
      jsonrpc: '2.0',
      result,
      id
    });
  } catch (error) {
    // Handle errors
    console.error('Error processing request:', error);
    
    // Format error for JSON-RPC response
    const errorResponse = {
      jsonrpc: '2.0',
      error: {
        code: error.code || -32603,
        message: error.message || 'Internal error'
      },
      id
    };
    
    return res.json(errorResponse);
  }
});

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`MCP Weather Tool server running on port ${PORT}`);
  console.log(`Debug output will appear below:`);
});
```

## Running the Server

1. Replace `YOUR_OPENWEATHERMAP_API_KEY` with your actual API key.

2. Start the server:

```bash
node server.js
```

You should see output confirming the server is running.

## Testing the Weather Tool

You can test your weather tool using curl:

```bash
curl -X POST http://localhost:3000 \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "method": "getCurrentWeather",
    "params": {
      "location": "London,UK",
      "units": "metric"
    },
    "id": 1
  }'
```

This should return current weather data for London in metric units.

## MCP Server Architecture

Our weather tool implements the MCP protocol with the following components:

1. **JSON-RPC Handling**: Processes incoming requests according to the JSON-RPC 2.0 specification
2. **Method Registration**: Exposes available methods during initialization
3. **Parameter Validation**: Ensures required parameters are provided
4. **External API Integration**: Communicates with OpenWeatherMap API
5. **Response Formatting**: Structures data in a consistent format

## Troubleshooting

### Common Issues

- **API Key Issues**: If you see `401 Unauthorized` errors, verify your API key is correct
- **Location Not Found**: Check that the location format is valid (e.g., "City,CountryCode")
- **Server Connection**: Ensure your server is running and accessible

### Debugging Tips

For additional debugging, you can add more console.log statements throughout the code. Focus on:

- Request payload validation
- API call parameters
- Response data from the weather API

## Next Steps

To enhance your weather tool, consider:

1. Adding forecasting capabilities
2. Implementing caching to reduce API calls
3. Adding geolocation support (coordinates instead of city names)
4. Supporting more detailed weather information

## Complete Code

The complete code for this tutorial is available in the [examples/weather-tool](../examples/weather-tool/) directory.
