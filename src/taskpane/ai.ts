import OpenAI from "openai";

export async function generate(baseURL: string, messages: OpenAI.Chat.Completions.ChatCompletionMessageParam[]) {
  console.log("Generating AI response");

  const openai = new OpenAI({
    baseURL,
    apiKey: "dummy-value",
    dangerouslyAllowBrowser: true,
    fetch: (input, init) => fetch(input, { ...(init ?? {}), headers: { "Content-Type": "application/json" } }),
  });

  const response = await openai.chat.completions.create({
    model: "gpt-3.5-turbo",
    messages,
    stream: false,
  });

  return response.choices[0].message.content;
}
